from __future__ import annotations

import re
import sys
import json
import logging
import tempfile
import textwrap
import time
import traceback
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats as scipy_stats

from openpyxl import load_workbook
from openpyxl.styles.fills import GradientFill, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from llm import call_llm

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)
from preprocessnew import load_data, detect_numeric_columns, _build_styles

#-----PDF GENERATION MODULE -------#

class PDFGenerationReport:
    """
    Generates a complete PDF data analysis report from a pandas DataFrame.

    This class handles dataset profiling, statistical analysis, LLM-based section
    selection, grounded narrative generation, and PDF rendering. All report content
    is based on Python-computed values so the LLM explains verified results instead
    of inventing or calculating data independently.

    Attributes:
        df (pd.DataFrame): Cleaned working DataFrame used for report generation.
    """

    SECTION_CATALOGUE = [
        {"id": "executive_summary", "title": "Executive Summary",
         "description": "High-level overview — key highlights, trends, and business takeaway. Best for any dataset.",
         "fn_name": "_gen_executive_summary"},
        {"id": "dataset_overview", "title": "Dataset Overview",
         "description": "Shape, column names, data types. Useful when dataset has many columns.",
         "fn_name": "_gen_dataset_overview"},
        {"id": "missing_values", "title": "Missing Value Analysis",
         "description": "Count and pattern of missing values per column.",
         "fn_name": "_gen_missing_values"},
        {"id": "statistical_summary", "title": "Statistical Summary",
         "description": "Mean, median, std, min, max, percentiles for all numeric columns.",
         "fn_name": "_gen_statistical_summary"},
        {"id": "trend_analysis", "title": "Trend Analysis",
         "description": "Average values and directional trends across numeric columns.",
         "fn_name": "_gen_trend_analysis"},
        {"id": "correlation_analysis", "title": "Correlation Analysis",
         "description": "Relationships between numeric variables. Most useful with 3+ numeric columns.",
         "fn_name": "_gen_correlation_analysis"},
        {"id": "outlier_detection", "title": "Anomaly & Outlier Detection",
         "description": "IQR- and Z-score-based outlier flags. Best for financial or operational data.",
         "fn_name": "_gen_outlier_detection"},
        {"id": "distribution_analysis", "title": "Distribution Analysis",
         "description": "Skewness, kurtosis, spread. Best for numeric-heavy datasets.",
         "fn_name": "_gen_distribution_analysis"},
        {"id": "segment_analysis", "title": "Segment / Categorical Analysis",
         "description": "Dominant categories, concentration, segment comparisons.",
         "fn_name": "_gen_categorical_analysis"},
        {"id": "kpi_analysis", "title": "KPI Analysis",
         "description": "Key performance indicators and their current levels.",
         "fn_name": "_gen_kpi_analysis"},
        {"id": "revenue_analysis", "title": "Revenue & Sales Analysis",
         "description": "Revenue trends and top performers. Only relevant with sales columns.",
         "fn_name": "_gen_revenue_analysis"},
        {"id": "cost_efficiency", "title": "Cost & Efficiency Analysis",
         "description": "Cost breakdown and efficiency ratios. Relevant with cost/profit columns.",
         "fn_name": "_gen_cost_efficiency"},
        {"id": "risk_assessment", "title": "Risk Assessment",
         "description": "Business and operational risks from data quality and distribution.",
         "fn_name": "_gen_risk_assessment"},
        {"id": "opportunities", "title": "Opportunities & Growth Areas",
         "description": "Untapped potential and quick wins.",
         "fn_name": "_gen_opportunities"},
        {"id": "forecasting", "title": "Forecasting & Future Trends",
         "description": "Scenario-based predictions from existing statistics.",
         "fn_name": "_gen_forecasting"},
        {"id": "geographic_analysis", "title": "Geographic & Regional Analysis",
         "description": "Performance by country/region. Only relevant with geographic columns.",
         "fn_name": "_gen_geographic_analysis"},
        {"id": "customer_analysis", "title": "Customer & Segment Behaviour",
         "description": "Customer patterns and churn risks. Best for CRM/sales datasets.",
         "fn_name": "_gen_customer_analysis"},
        {"id": "strategic_recommendations", "title": "Strategic Recommendations",
         "description": "Prioritised, actionable recommendations for leadership.",
         "fn_name": "_gen_strategic_recommendations"},
        {"id": "key_questions", "title": "Key Strategic Questions",
         "description": "20 sharp questions a C-suite executive should ask.",
         "fn_name": "_gen_key_questions"},
    ]

    SELECTOR_SYSTEM = textwrap.dedent("""\
        You are an expert report architect.
        You will be given a structured profile of a dataset and a catalogue of possible
        report sections. Select exactly 10 most relevant and valuable sections.

        Rules:
        - Return ONLY a valid JSON array of exactly 10 section IDs.
        - Choose sections genuinely applicable to this data.
        - Avoid sections clearly irrelevant (e.g. geographic_analysis if no location columns).
        - Always include executive_summary, strategic_recommendations, and key_questions.
        - Return NO prose, NO markdown, NO explanation — only the JSON array.
    """)

    MONO_SECTIONS = {"dataset_overview", "missing_values", "statistical_summary"}

    META_COLS = ["__sheet__", "__table__", "__excel_name__", "__range__"]

    def __init__(self, df: pd.DataFrame) -> None:
        """
        Initializes the PDF report generator.

        The raw DataFrame (including metadata columns __sheet__, __table__,
        __excel_name__, and __range__) is split into per-table groups so that
        each table receives its own PDF section with 10 subsections.
        Metadata columns are removed from the analytical DataFrame passed to
        each section generator.

        Args:
            df (pd.DataFrame): Input DataFrame loaded from the source file.

        Returns:
            None
        """
        self._raw_df = df  # full df with meta columns intact

        # Build list of (table_meta_dict, clean_df) pairs – one entry per table
        self._tables: List[Tuple[dict, pd.DataFrame]] = []
        group_cols = [c for c in ["__sheet__", "__table__"] if c in df.columns]
        if group_cols:
            for keys, grp in df.groupby(group_cols, sort=False):
                if not isinstance(keys, tuple):
                    keys = (keys,)
                meta = {
                    "sheet_name": str(keys[0]) if len(keys) > 0 else "",
                    "table_num": int(keys[1]) if len(keys) > 1 else 1,
                    "excel_name": str(grp["__excel_name__"].iloc[0]) if "__excel_name__" in grp.columns else "",
                    "range": str(grp["__range__"].iloc[0]) if "__range__" in grp.columns else "",
                }
                clean = grp.drop(columns=self.META_COLS, errors="ignore").reset_index(drop=True)
                self._tables.append((meta, clean))
        else:
            # No grouping columns – treat entire df as one table
            meta = {"sheet_name": "", "table_num": 1, "excel_name": "", "range": ""}
            clean = df.drop(columns=self.META_COLS, errors="ignore")
            self._tables.append((meta, clean))

        # Combined clean df kept for helpers that still need it
        self.df = df.drop(columns=self.META_COLS, errors="ignore")

    def run(self, output: str = "report.pdf") -> None:
        """
        Runs the full PDF report generation workflow.

        For every detected table the method:
          1. Runs statistical analysis on that table's DataFrame.
          2. Asks the LLM to select exactly 10 subsections for that table.
          3. Generates content for each subsection.
          4. Collects the table metadata (sheet name, excel name, range) so it
             can be printed just below the table heading in the PDF.

        All per-table results are handed to build_pdf() which renders one
        top-level section per table, each containing 10 numbered subsections.

        Args:
            output (str): Output path where the generated PDF file will be saved.

        Returns:
            None
        """
        catalogue_map = {s["id"]: s for s in self.SECTION_CATALOGUE}

        # per_table_data: list of (meta, analysis, sections_output)
        per_table_data: List[Tuple[dict, dict, List[Tuple[str, str, str]]]] = []

        with tempfile.TemporaryDirectory():
            for t_idx, (meta, df_table) in enumerate(self._tables):
                table_label = f"Table {meta['table_num']} (sheet: {meta['sheet_name']})"
                log.info("=== Processing %s (%d rows) ===", table_label, len(df_table))

                if df_table.empty:
                    log.warning("  Table is empty, skipping.")
                    continue

                log.info("  Running statistical analysis...")
                analysis = self.analyze_data(df_table)
                log.info(
                    "  Analysis: %d numeric, %d categorical, %d strong corr, %d outlier cols",
                    len(analysis["numeric_columns"]),
                    len(analysis["categorical_columns"]),
                    len(analysis["strong_correlations"]),
                    sum(1 for s in analysis["column_stats"].values() if s["iqr_outlier_count"] > 0),
                )

                log.info("  Asking LLM to select 10 subsections...")
                selected_ids = self.select_sections(analysis)
                log.info("  LLM selected: %s", selected_ids)

                log.info("  Generating LLM context...")
                llm_context = self.generate_llm_context(df_table)

                log.info("  Generating content for %d subsections...", len(selected_ids))
                sections_output: List[Tuple[str, str, str]] = []
                for sid in selected_ids:
                    if sid not in catalogue_map:
                        log.warning("  Unknown section id '%s', skipping.", sid)
                        continue

                    entry = catalogue_map[sid]
                    log.info("    Generating: %s", entry["title"])

                    try:
                        fn = getattr(self, entry["fn_name"])
                        content = fn({**analysis, "llm_context": llm_context})
                    except Exception:
                        log.warning("    Subsection '%s' failed:\n%s", sid, traceback.format_exc())
                        content = "[Subsection generation failed — see logs]"

                    if not content or not content.strip():
                        content = "[No content generated for this subsection]"

                    sections_output.append((sid, entry["title"], content))

                per_table_data.append((meta, analysis, sections_output))

        self.build_pdf(output, per_table_data)

        print(f"\n  Report saved : {output}")
        print(f"  Tables       : {len(per_table_data)}")
        for t_idx, (meta, _, secs) in enumerate(per_table_data, 1):
            print(f"\n  Table {t_idx}: {meta['sheet_name']} | {meta['excel_name']} | {meta['range']}")
            for i, (sid, title, _) in enumerate(secs, 1):
                print(f"    {i:2}. {title}")

    def clean_llm_output(self, text: str) -> str:
        """
        LLM responses may include internal text before a `</think>` marker.
        This function keeps only the final response after that marker so the report
        contains clean, user-facing content.

        Args:
            text (str): Raw text returned by the LLM.

        Returns:
            str: Cleaned LLM response text.
        """

        parts = text.split("</think>", 1)
        return parts[1].strip() if len(parts) > 1 else text


    def safe_llm(self, system: str, user: str, fallback: str = "[Content unavailable]") -> str:
        """
        Calls the nemotron  safely and normalizes the returned text.

        Args:
            system (str): System prompt that defines the LLM role and rules.
            user (str): User prompt containing the task and grounded data.
            fallback (str): Text returned if the LLM call fails.

        Returns:
            str: Cleaned and PDF-safe LLM output, or the fallback message.
        """

        try:
            result = call_llm(system, user)

            result = self.clean_llm_output(result)

            return result.encode("latin-1", errors="replace").decode("latin-1")
        except Exception as exc:
            log.warning("LLM call failed: %s", exc)
            return fallback

    def llm_narrative(self, role: str, task: str, grounded_data: dict) -> str:
        """
        Generates an LLM-written narrative for a report section.

        The function converts Python-computed data into JSON, builds a strict prompt
        to prevent hallucination, and sends the task plus grounded data to safe_llm().

        Args:
            role (str): Analyst role the LLM should follow for the narrative.
            task (str): Specific instruction for the report section.
            grounded_data (dict): Python-computed facts the LLM is allowed to use.

        Returns:
            str: LLM-generated narrative grounded only in the provided data.
        """
        context_json = json.dumps(grounded_data, indent=2, default=str)
        
        system = f"""
                    You are a {role}.

                    STRICT RULES:
                    - Use ONLY provided data
                    - DO NOT invent or infer missing values
                    """

        user = f"{task}\n\nPRE-COMPUTED DATA (Python-verified):\n{context_json}"
        return self.safe_llm(system, user)

    @staticmethod
    def _safe_float(x) -> Optional[float]:
        """
        Safely converts a value into a rounded float for reporting.

        If the value cannot be converted, or if it becomes NaN or infinity,
        the function returns None so invalid numbers do not appear in the PDF.

        Args:
            x: Value to convert into a float.

        Returns:
            Optional[float]: Rounded float value, or None if conversion is invalid.
        """

        try:
            v = float(x)
            return None if (np.isnan(v) or np.isinf(v)) else round(v, 6)
        except Exception:
            return None

    def analyze_data(self, df: pd.DataFrame) -> dict:
        """
        Performs the main statistical analysis on the extracted dataset.

        This function identifies numeric and categorical columns, cleans numeric
        values, calculates descriptive statistics, detects missing values,
        finds duplicate rows, computes correlations, and flags outliers using
        IQR and Z-score methods.

        Args:
            df (pd.DataFrame): Input dataset to analyze.

        Returns:
            dict: Python-computed analysis results used by PDF sections and LLM
            narratives, including shape, columns, dtypes, statistics, correlations,
            missing-value summary, duplicate count, and outlier details.
        """


        #editing copy
        df  = df.copy()
        numeric_cols = detect_numeric_columns(df)
        for col in numeric_cols:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("$", "", regex=False)
                .str.replace("₹", "", regex=False)
                .str.replace("%", "", regex=False)
                .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")
        numeric_df = df[numeric_cols]

        cat_df = df.drop(columns=numeric_cols, errors="ignore")
        # cat_df = df.select_dtypes(include=["object", "category"])
        print(numeric_df, cat_df)
        n_rows, n_cols = df.shape

        column_stats: Dict[str, dict] = {}
        for col in numeric_df.columns:
            s = numeric_df[col].dropna()
            if s.empty:
                continue

            if len(s) < 2:
                skewness = None
                kurtosis = None
            else:
                skewness = self._safe_float(scipy_stats.skew(s))
                kurtosis = self._safe_float(scipy_stats.kurtosis(s))

            q1, q3 = float(s.quantile(0.25)), float(s.quantile(0.75))
            iqr = q3 - q1
            lower_fence = q1 - 1.5 * iqr
            upper_fence = q3 + 1.5 * iqr
            outlier_mask = (s < lower_fence) | (s > upper_fence)
            outlier_vals = s[outlier_mask].tolist()[:10]

            if len(s) >= 2:
                z_scores = np.abs(scipy_stats.zscore(s))
                z_outlier_n = int((z_scores > 3).sum())
            else:
                z_outlier_n = 0

            skew_val = skewness or 0
            column_stats[col] = {
                "count": int(s.count()),
                "missing": int(df[col].isnull().sum()),
                "missing_pct": round(df[col].isnull().mean() * 100, 2),
                "mean": self._safe_float(s.mean()),
                "median": self._safe_float(s.median()),
                "std": self._safe_float(s.std()),
                "variance": self._safe_float(s.var()),
                "min": self._safe_float(s.min()),
                "max": self._safe_float(s.max()),
                "range": self._safe_float(s.max() - s.min()),
                "p05": self._safe_float(s.quantile(0.05)),
                "p25": self._safe_float(q1),
                "p75": self._safe_float(q3),
                "p95": self._safe_float(s.quantile(0.95)),
                "iqr": self._safe_float(iqr),
                "skewness": skewness,
                "kurtosis": kurtosis,
                "skew_label": (
                    "right-skewed" if skew_val > 0.5
                    else "left-skewed" if skew_val < -0.5
                    else "symmetric"
                ),
                "iqr_outlier_count": int(outlier_mask.sum()),
                "iqr_outlier_pct": round(outlier_mask.mean() * 100, 2),
                "iqr_outlier_examples": [round(v, 4) for v in outlier_vals],
                "zscore_outlier_count": z_outlier_n,
                "lower_fence": self._safe_float(lower_fence),
                "upper_fence": self._safe_float(upper_fence),
            }

        correlations: Dict[str, float] = {}
        strong_correlations: List[dict] = []
        if numeric_df.shape[1] >= 2:
            corr_matrix = numeric_df.corr(method="pearson")
            cols = corr_matrix.columns.tolist()
            for i, c1 in enumerate(cols):
                for c2 in cols[i + 1:]:
                    r = self._safe_float(corr_matrix.loc[c1, c2])
                    if r is not None:
                        key = f"{c1} / {c2}"
                        correlations[key] = r
                        if abs(r) >= 0.7:
                            strong_correlations.append({
                                "col_a": c1, "col_b": c2,
                                "r": r,
                                "strength": "very strong" if abs(r) >= 0.9 else "strong",
                                "direction": "positive" if r > 0 else "negative",
                            })
        strong_correlations.sort(key=lambda x: abs(x["r"]), reverse=True)

        categorical_stats: Dict[str, dict] = {}
        for col in cat_df.columns:
            vc = df[col].value_counts()
            categorical_stats[col] = {
                "unique_count": int(df[col].nunique()),
                "missing": int(df[col].isnull().sum()),
                "missing_pct": round(df[col].isnull().mean() * 100, 2),
                "top_5_values": {str(k): int(v) for k, v in vc.head(5).items()},
                "top_value": str(vc.index[0]) if not vc.empty else None,
                "top_value_pct": round(vc.iloc[0] / len(df) * 100, 2) if not vc.empty else None,
            }

        total_missing = int(df.isnull().sum().sum())
        total_cells = n_rows * n_cols
        cols_with_missing = [c for c in df.columns if df[c].isnull().any()]
        high_missing_cols = [c for c in df.columns if df[c].isnull().mean() > 0.20]
        duplicate_rows = int(df.duplicated().sum())

        numeric_means = {
            col: column_stats[col]["mean"]
            for col in column_stats
            if column_stats[col]["mean"] is not None
        }

        return {
            "shape": {"rows": n_rows, "cols": n_cols},
            "columns": df.columns.tolist(),
            "numeric_columns": numeric_df.columns.tolist(),
            "categorical_columns": cat_df.columns.tolist(),
            "dtypes": {c: str(t) for c, t in df.dtypes.items()},
            "column_stats": column_stats,
            "categorical_stats": categorical_stats,
            "correlations": correlations,
            "strong_correlations": strong_correlations,
            "numeric_means": numeric_means,
            "missing_summary": {
                "total_missing": total_missing,
                "total_cells": total_cells,
                "missing_pct": round(total_missing / total_cells * 100, 2) if total_cells else 0,
                "cols_with_missing": cols_with_missing,
                "high_missing_cols": high_missing_cols,
            },
            "duplicate_rows": duplicate_rows,
        }

    def generate_llm_context(self, df: pd.DataFrame) -> dict:
        """
        Builds a compact, structured summary of the dataset for LLM consumption.

        Instead of sending the full dataframe to the LLM, this function prepares
        key context such as column types, sample rows, missing values, top category
        values, numeric summaries, and likely column roles.

        Args:
            df (pd.DataFrame): Dataset used to build summarized LLM context.

        Returns:
            dict: Compact dataset profile containing row count, column count, column
            types, cardinality, sample rows, missing values, duplicate rows, numeric
            summaries, and inferred column roles.
        """


        n = len(df)
        context: dict = {
            "row_count": n,
            "column_count": len(df.columns),
        }

        types: Dict[str, str] = {}
        numeric_cols: List[str] = []
        categorical_cols: List[str] = []
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                types[col] = "numeric"
                numeric_cols.append(col)
            else:
                types[col] = "categorical"
                categorical_cols.append(col)
        context["types"] = types
        context["numeric_columns"] = numeric_cols
        context["categorical_columns"] = categorical_cols

        cardinality: Dict[str, str] = {}
        for col in df.columns:
            unique_ratio = df[col].nunique() / n if n > 0 else 0
            if unique_ratio > 0.9:
                cardinality[col] = "high"
            elif unique_ratio > 0.3:
                cardinality[col] = "medium"
            else:
                cardinality[col] = "low"
        context["cardinality"] = cardinality

        top_values: Dict[str, dict] = {}
        for col in categorical_cols:
            vc = df[col].value_counts().head(5)
            top_values[col] = vc.to_dict()
        context["top_values"] = top_values

        context["sample_rows"] = df.head(5).to_dict(orient="records")
        context["missing_values"] = df.isnull().sum().to_dict()
        context["duplicate_rows"] = int(df.duplicated().sum())

        numeric_stats: Dict[str, dict] = {}
        for col in numeric_cols:
            numeric_stats[col] = {
                "mean": float(df[col].mean()),
                "min": float(df[col].min()),
                "max": float(df[col].max()),
            }
        context["numeric_stats"] = numeric_stats

        identifier_cols: List[str] = []
        groupable_cols: List[str] = []
        text_cols: List[str] = []
        for col in df.columns:
            if cardinality[col] == "high":
                identifier_cols.append(col)
            elif cardinality[col] == "low":
                groupable_cols.append(col)
            if df[col].dtype == "object":
                text_cols.append(col)
        context["column_roles"] = {
            "identifier_columns": identifier_cols,
            "groupable_columns": groupable_cols,
            "text_columns": text_cols,
        }
        return context

    def select_sections(self, analysis: dict) -> List[str]:
        """
        Selects the most relevant report sections based on the dataset profile.

        This function sends a summarized analysis profile and the available section
        catalogue to the LLM, asking it to choose exactly 10 sections that best fit
        the data. If the LLM response is invalid or cannot be parsed, the function
        falls back to a default set of general-purpose analytical sections.

        Args:
            analysis (dict): Python-computed dataset analysis dictionary.

        Returns:
            List[str]: Exactly 10 selected section IDs.
        """

        catalogue_text = "\n".join(
            f'  id: "{s["id"]}"  |  when_useful: "{s["description"]}"'
            for s in self.SECTION_CATALOGUE
        )
        profile = {
            "shape": analysis["shape"],
            "numeric_columns": analysis["numeric_columns"],
            "categorical_columns": analysis["categorical_columns"],
            "columns_with_missing": analysis["missing_summary"]["cols_with_missing"],
            "has_strong_correlations": len(analysis["strong_correlations"]) > 0,
            "has_outliers": any(
                s["iqr_outlier_count"] > 0 for s in analysis["column_stats"].values()
            ),
        }
        user_prompt = (
            f"DATASET PROFILE:\n{json.dumps(profile, indent=2)}\n\n"
            f"SECTION CATALOGUE:\n{catalogue_text}\n\n"
            "Select exactly 10 section IDs most relevant to this dataset. "
            "Return only a JSON array of 10 IDs."
        )
        raw = self.safe_llm(self.SELECTOR_SYSTEM, user_prompt)
        log.info("LLM section selector raw: %s", raw[:300])

        defaults = [
            "executive_summary", "dataset_overview", "statistical_summary",
            "trend_analysis", "correlation_analysis", "outlier_detection",
            "distribution_analysis", "kpi_analysis", "strategic_recommendations", "key_questions",
        ]
        MANDATORY = {"executive_summary", "strategic_recommendations", "key_questions"}
        known = {s["id"] for s in self.SECTION_CATALOGUE}
        match = re.search(r"\[.*?\]", raw, re.DOTALL)

        if not match:
            log.warning("Could not parse section IDs from LLM. Using defaults.")
            return defaults

        try:
            ids = json.loads(match.group())
            ids = list(dict.fromkeys(i for i in ids if i in known))

            # Ensure all mandatory sections are present
            for m_id in MANDATORY:
                if m_id not in ids:
                    ids.insert(0, m_id)

            # Trim to 10 while preserving mandatory sections
            if len(ids) > 10:
                non_mandatory = [i for i in ids if i not in MANDATORY]
                mandatory_in_ids = [i for i in ids if i in MANDATORY]
                ids = mandatory_in_ids + non_mandatory[:10 - len(mandatory_in_ids)]

            # Pad to exactly 10 using defaults if needed
            while len(ids) < 10:
                for d in defaults:
                    if d not in ids:
                        ids.append(d)
                    if len(ids) == 10:
                        break

            log.info("Selected sections: %s", ids)
            return ids
        except json.JSONDecodeError:
            log.warning("JSON parse error on section IDs. Using defaults.")
            return defaults

    def _gen_executive_summary(self, a: dict) -> str:
        """
        Generates the executive summary section of the PDF report.

        This function prepares a compact set of high-level dataset insights such as
        dataset size, missing data, duplicate rows, strong correlations, numeric means,
        and top outlier columns, then passes that verified context to the LLM.

        Args:
            a (dict): Analysis dictionary containing dataset metadata, missing values,
            duplicate counts, correlations, numeric means, and column statistics.

        Returns:
            str: Executive summary narrative generated from verified analysis data.
        """
        grounded = {
            "shape": a["shape"],
            "columns": a["columns"],
            "missing_summary": a["missing_summary"],
            "duplicate_rows": a["duplicate_rows"],
            "strong_correlations": a["strong_correlations"][:5],
            "numeric_means": a["numeric_means"],
            "top_outlier_cols": [
                {"col": c, "outlier_count": v["iqr_outlier_count"], "outlier_pct": v["iqr_outlier_pct"]}
                for c, v in a["column_stats"].items()
                if v["iqr_outlier_count"] > 0
            ][:5],
        }
        task = """
            Write an executive summary using ONLY the provided data.

            Requirements:
            - Maximum 100 words
            - Must include:
            1. Dataset size (rows, columns)
            2. Missing data (count and percentage)
            3. Correlations (only if present)
            4. One clear business takeaway

        """
        return self.llm_narrative(
    "senior business analyst",
    task,
    grounded,
)

    def _gen_dataset_overview(self, a: dict) -> str:
        """
        Generates the dataset overview section of the PDF report.

        This function creates a technical summary of the dataset, including the
        total number of rows, total number of columns, column names, and detected
        data types.

        Args:
            a (dict): Analysis dictionary containing shape, columns, and dtypes.

        Returns:
            str: Plain-text dataset overview formatted for monospace PDF output.
        """
   
        lines = [
            f"Rows:    {a['shape']['rows']:,}",
            f"Columns: {a['shape']['cols']}",
            "",
            "Column list:",
        ]
        for col in a["columns"]:
            dtype = a["dtypes"].get(col, "unknown")
            lines.append(f"  {col:<35} {dtype}")
        return "\n".join(lines)

    def _gen_missing_values(self, a: dict) -> str:
        """
        Generates the missing value analysis section of the PDF report.

        This function summarizes overall missing data, duplicate rows, and
        per-column missing counts using the Python-computed analysis dictionary.
        It also highlights high-risk columns where more than 20% of values are
        missing.

        Args:
            a (dict): Analysis dictionary containing missing summary, duplicate
            count, column statistics, categorical statistics, and column names.

        Returns:
            str: Plain-text missing-value summary formatted for PDF output.
        """


        ms = a["missing_summary"]
        lines = [
            f"Total missing cells : {ms['total_missing']:,}  ({ms['missing_pct']} % of all cells)",
            f"Duplicate rows      : {a['duplicate_rows']:,}",
            "",
            "Per-column missing counts:",
        ]
        for col in a["columns"]:
            if col in a["column_stats"]:
                s = a["column_stats"][col]
                lines.append(f"  {col:<35} {s['missing']:>6,}  ({s['missing_pct']} %)")
            elif col in a["categorical_stats"]:
                s = a["categorical_stats"][col]
                lines.append(f"  {col:<35} {s['missing']:>6,}  ({s['missing_pct']} %)")
        if ms["high_missing_cols"]:
            lines += ["", "Columns with > 20 % missing (high-risk):"]
            lines += [f"  {c}" for c in ms["high_missing_cols"]]
        return "\n".join(lines)

    def _gen_statistical_summary(self, a: dict) -> str:
        """
        Generates the statistical summary section of the PDF report.

        This function formats key descriptive statistics for each numeric column,
        including count, mean, median, standard deviation, minimum, maximum,
        and skewness.

        Args:
            a (dict): Analysis dictionary containing computed numeric column statistics.

        Returns:
            str: Aligned plain-text statistical summary for monospace PDF output.
        """

        rows = [
            f"{'Column':<30} {'Count':>7} {'Mean':>14} {'Median':>14} "
            f"{'Std':>14} {'Min':>14} {'Max':>14} {'Skew':>8}"
        ]
        rows.append("-" * 113)
        for col, s in a["column_stats"].items():
            def f(v):
                """
                Formats a numeric value for the statistical summary table.

                Args:
                    v: Numeric value to format.

                Returns:
                    str: Right-aligned formatted number, or N/A if the value is missing.
                """
                return f"{v:>14,.4f}" if v is not None else f"{'N/A':>14}"
            rows.append(
                f"{col:<30} {s['count']:>7,} {f(s['mean'])} {f(s['median'])} "
                f"{f(s['std'])} {f(s['min'])} {f(s['max'])} "
                f"{str(s['skewness'] or 'N/A'):>8}"
            )
        return "\n".join(rows)

    def _gen_trend_analysis(self, a: dict) -> str:
        """
        Generates the trend analysis section of the PDF report.

        This function prepares key numeric statistics such as mean, median,
        standard deviation, skew label, and percentile values for each numeric
        column, then sends them to the LLM for interpretation.

        Args:
            a (dict): Analysis dictionary containing numeric means and column statistics.

        Returns:
            str: Trend-style narrative generated from verified numeric summaries.
        """

        grounded = {
            "numeric_means": a["numeric_means"],
            "column_details": {
                col: {k: v for k, v in a["column_stats"][col].items()
                      if k in ("mean", "median", "std", "skew_label", "p25", "p75")}
                for col in a["column_stats"]
            },
        }
        task = """
                Analyze numeric columns using provided statistics.

                Output:
                - Highest mean column
                - Lowest mean column
                - Most variable column (highest std)
                - Any skewed columns (left/right)

             """
        return self.llm_narrative(
            "data analyst",
            task,
            grounded,
        )

    def _gen_correlation_analysis(self, a: dict) -> str:
        """
        Generates the correlation analysis section of the PDF report.

        This function uses Python-computed Pearson correlation values to identify
        relationships between numeric columns. Strong correlations are already
        filtered during analysis using the configured threshold.

        Args:
            a (dict): Analysis dictionary containing all correlations and strong
            correlation pairs.

        Returns:
            str: Correlation interpretation, or a message if correlation analysis
            is not possible.
        """


        if not a["correlations"]:
            return "Not enough numeric columns for correlation analysis."
        grounded = {
            "strong_correlations": a["strong_correlations"],
            "all_correlations": a["correlations"],
        }
        return self.llm_narrative(
            "data scientist",
            "Interpret the correlation values. For each strong pair (|r| >= 0.7), "
            "explain the likely relationship and its business implication. "
            "Use the exact r-values provided — do not round or alter them.",
            grounded,
        )

    def _gen_outlier_detection(self, a: dict) -> str:
        """
        Generates the anomaly and outlier detection section of the PDF report.

        This function collects numeric columns where Python detected outliers using
        either the IQR method or the Z-score method, along with counts, percentages,
        example values, and boundary thresholds.

        Args:
            a (dict): Analysis dictionary containing numeric column statistics and
            outlier details.

        Returns:
            str: Outlier interpretation, or a message if no outliers are detected.
        """


        outlier_summary = [
            {
                "column": col,
                "iqr_outliers": s["iqr_outlier_count"],
                "iqr_pct": s["iqr_outlier_pct"],
                "z_outliers": s["zscore_outlier_count"],
                "examples": s["iqr_outlier_examples"],
                "lower_fence": s["lower_fence"],
                "upper_fence": s["upper_fence"],
            }
            for col, s in a["column_stats"].items()
            if s["iqr_outlier_count"] > 0 or s["zscore_outlier_count"] > 0
        ]
        if not outlier_summary:
            return "No outliers detected by IQR or Z-score methods across any numeric column."
        return self.llm_narrative(
            "risk analyst",
            "Interpret these outlier findings. For each flagged column, explain whether "
            "the outliers represent data quality issues, genuine business anomalies, or "
            "expected extremes. Use the exact counts and example values provided.",
            {"outlier_summary": outlier_summary},
        )

    def _gen_distribution_analysis(self, a: dict) -> str:
        """
        Generates the distribution analysis section of the PDF report.

        This function prepares distribution-related statistics for each numeric
        column, including mean, median, standard deviation, skewness, kurtosis,
        and percentile range.

        Args:
            a (dict): Analysis dictionary containing distribution statistics for
            numeric columns.

        Returns:
            str: Distribution interpretation generated from verified statistics.
        """



        dist_data = {
            col: {
                "mean": s["mean"], "median": s["median"], "std": s["std"],
                "skewness": s["skewness"], "kurtosis": s["kurtosis"],
                "skew_label": s["skew_label"], "p05": s["p05"], "p95": s["p95"],
            }
            for col, s in a["column_stats"].items()
        }
        task = """
                Describe distribution for each numeric column.

                Output for each column:
                [Column] — symmetric / left-skewed / right-skewed — implication (1 line)

            """
        return self.llm_narrative(
            "statistician",
            task, 
            {"distributions": dist_data},
        )

    def _gen_categorical_analysis(self, a: dict) -> str:

        """
        Generates the categorical or segment analysis section of the PDF report.

        This function checks whether the dataset contains categorical columns and,
        if available, passes their Python-computed value distributions to the LLM.

        Args:
            a (dict): Analysis dictionary containing categorical column statistics.

        Returns:
            str: Categorical analysis narrative, or a message if no categorical
            columns are available.
        """


        if not a["categorical_stats"]:
            return "No categorical columns found in this dataset."
        return self.llm_narrative(
            "market analyst",
            "Analyse the categorical column distributions. Identify dominant categories, "
            "potential concentration risks, and segment-level insights.",
            {"categorical_stats": a["categorical_stats"]},
        )

    def _gen_kpi_analysis(self, a: dict) -> str:
        """
        Generates the KPI analysis section of the PDF report.

        This function uses available column names, numeric means, and selected
        descriptive statistics to identify possible KPI-style metrics. The LLM
        is instructed to evaluate only the provided data and classify each metric
        as high, medium, or low performance with any visible concerns.

        Args:
            a (dict): Analysis dictionary containing dataset columns, numeric means,
            and computed statistics for numeric columns.

        Returns:
            str: KPI analysis narrative generated from Python-computed statistics.
        """
        grounded = {
            "columns": a["columns"],
            "numeric_means": a["numeric_means"],
            "column_stats": {
                col: {k: v for k, v in s.items()
                      if k in ("mean", "median", "min", "max", "std", "iqr_outlier_pct")}
                for col, s in a["column_stats"].items()
            },
        }
        task = """
            Identify KPIs strictly from provided column names and values.

            Output format:
            [Column Name] — performance level (high/medium/low) — concern (if any)


        """
        return self.llm_narrative(
            "business performance expert",
            task, 
            grounded,
        )

    def _gen_revenue_analysis(self, a: dict) -> str:
        """
        Generates the revenue and sales analysis section of the PDF report.

        This function identifies revenue-related numeric columns by checking column
        names for business keywords such as revenue, sales, profit, income, gross,
        net, and amount.

        Args:
            a (dict): Analysis dictionary containing numeric columns, column statistics,
            and strong correlation details.

        Returns:
            str: Revenue analysis narrative, or a message if no revenue-like columns
            are detected.
        """

        revenue_cols = [
            c for c in a["numeric_columns"]
            if any(k in c.lower() for k in ("revenue", "sales", "gross", "net", "profit", "income", "amount"))
        ]
        if not revenue_cols:
            return "No columns clearly identifiable as revenue or sales were found. Columns present: " + ", ".join(a["numeric_columns"])
        grounded = {
            "revenue_columns": {col: a["column_stats"][col] for col in revenue_cols if col in a["column_stats"]},
            "strong_correlations": [
                x for x in a["strong_correlations"]
                if x["col_a"] in revenue_cols or x["col_b"] in revenue_cols
            ],
        }
        return self.llm_narrative(
            "revenue analyst",
            "Analyse revenue and sales patterns using these exact statistics. "
            "Comment on magnitude, variability, and any correlations with other columns.",
            grounded,
        )

    def _gen_cost_efficiency(self, a: dict) -> str:
        """
        Generates the cost and efficiency analysis section of the PDF report.

        This function identifies cost-related numeric columns by checking column
        names for keywords such as cost, expense, COGS, margin, discount, and spend.

        Args:
            a (dict): Analysis dictionary containing numeric columns and column
            statistics.

        Returns:
            str: Cost and efficiency analysis narrative, or a message if no relevant
            columns are detected.
        """

        cost_cols = [
            c for c in a["numeric_columns"]
            if any(k in c.lower() for k in ("cost", "expense", "cogs", "margin", "discount", "spend"))
        ]
        if not cost_cols:
            return "No columns clearly identifiable as cost or efficiency metrics were found. Columns present: " + ", ".join(a["numeric_columns"])
        grounded = {"cost_columns": {col: a["column_stats"][col] for col in cost_cols if col in a["column_stats"]}}
        return self.llm_narrative(
            "financial analyst",
            "Analyse cost and efficiency patterns. Comment on variability, outliers, "
            "and what these numbers suggest about operational efficiency.",
            grounded,
        )

    def _gen_risk_assessment(self, a: dict) -> str:
        """
        Generates the risk assessment section of the PDF report.

        This function prepares risk-related indicators such as missing data,
        duplicate rows, detected outliers, and high-variation numeric columns.

        Args:
            a (dict): Analysis dictionary containing missing values, duplicate rows,
            outlier statistics, and numeric column statistics.

        Returns:
            str: Risk assessment narrative generated from verified risk indicators.
        """


        grounded = {
            "missing_summary": a["missing_summary"],
            "duplicate_rows": a["duplicate_rows"],
            "outlier_summary": [
                {"column": col, "iqr_outliers": s["iqr_outlier_count"], "pct": s["iqr_outlier_pct"]}
                for col, s in a["column_stats"].items() if s["iqr_outlier_count"] > 0
            ],
            "high_dispersion_cols": [
                {"column": col, "std": s["std"], "mean": s["mean"],
                 "cv": round((s["std"] or 0) / (s["mean"] or 1), 4)}
                for col, s in a["column_stats"].items()
                if s["mean"] and s["std"] and abs(s["std"] / s["mean"]) > 1
            ],
        }
        return self.llm_narrative(
            "risk management consultant",
            "Assess business and operational risks visible in this data. "
            "Rate each risk (High / Medium / Low) and explain it using the exact figures provided.",
            grounded,
        )

    def _gen_opportunities(self, a: dict) -> str:
        """
        Generates the opportunities and growth areas section of the PDF report.

        This function prepares high-level opportunity signals such as numeric
        averages, strong correlations, available columns, and dataset size.

        Args:
            a (dict): Analysis dictionary containing numeric means, strong correlations,
            columns, and dataset shape.

        Returns:
            str: Opportunities narrative grounded in verified dataset statistics.
        """
        grounded = {
            "numeric_means": a["numeric_means"],
            "strong_correlations": a["strong_correlations"][:5],
            "columns": a["columns"],
            "shape": a["shape"],
        }
        return self.llm_narrative(
            "growth strategist",
            "Identify the top 5 opportunities or high-potential growth areas evident from "
            "this data. Ground each opportunity in the actual statistics provided.",
            grounded,
        )

    def _gen_forecasting(self, a: dict) -> str:
        """
        Generates the forecasting and future trends section of the PDF report.

        This function prepares basic forward-looking indicators such as numeric
        means, standard deviations, skew labels, and strong correlations. It does
        not perform true time-series forecasting.

        Args:
            a (dict): Analysis dictionary containing numeric means, column statistics,
            and strong correlations.

        Returns:
            str: Cautious directional outlook based only on available summary statistics.
        """

        grounded = {
            "numeric_means": a["numeric_means"],
            "column_stats": {
                col: {"mean": s["mean"], "std": s["std"], "skew_label": s["skew_label"]}
                for col, s in a["column_stats"].items()
            },
            "strong_correlations": a["strong_correlations"][:3],
        }
        task = """
            Provide simple directional outlook based on means and variability.

            Output:
            - Upward / stable / uncertain trend for each key metric

            Rules:
            - Base ONLY on mean and std
            - DO NOT predict exact numbers
            - DO NOT assume time-series

            If insufficient data:
            "Insufficient data for forecasting"
        """
        return self.llm_narrative(
            "forecasting expert",
            task, 
            grounded,
        )

    def _gen_geographic_analysis(self, a: dict) -> str:
        """
        Generates the geographic and regional analysis section of the PDF report.

        This function identifies location-related categorical columns by checking
        column names for keywords such as country, region, city, state, location,
        geography, geo, and market.

        Args:
            a (dict): Analysis dictionary containing categorical columns and
            categorical statistics.

        Returns:
            str: Geographic analysis narrative, or a message if no geographic
            columns are detected.
        """

        geo_cols = [
            c for c in a["categorical_columns"]
            if any(k in c.lower() for k in ("country", "region", "city", "state", "location", "geography", "geo", "market"))
        ]
        if not geo_cols:
            return "No geographic or regional columns detected in this dataset."
        grounded = {"geographic_columns": {col: a["categorical_stats"][col] for col in geo_cols if col in a["categorical_stats"]}}
        return self.llm_narrative(
            "regional business analyst",
            "Analyse geographic or regional performance differences using the provided "
            "category distributions. Identify dominant regions and any concentration risks.",
            grounded,
        )

    def _gen_customer_analysis(self, a: dict) -> str:
        """
        Generates the customer and segment behaviour section of the PDF report.

        This function identifies customer-related categorical columns by checking
        column names for keywords such as customer, segment, channel, product,
        and category.

        Args:
            a (dict): Analysis dictionary containing categorical columns, categorical
            statistics, numeric means, and strong correlations.

        Returns:
            str: Customer or segment behaviour narrative generated from verified data.
        """

        cust_cols = [
            c for c in a["categorical_columns"]
            if any(k in c.lower() for k in ("customer", "segment", "channel", "product", "category"))
        ]
        grounded = {
            "categorical_stats": {col: a["categorical_stats"][col] for col in cust_cols if col in a["categorical_stats"]},
            "numeric_means": a["numeric_means"],
            "strong_correlations": a["strong_correlations"][:3],
        }
        return self.llm_narrative(
            "customer insights analyst",
            "Identify customer or segment behaviour patterns, high-value segments, "
            "and churn or concentration risks using the provided statistics.",
            grounded,
        )

    def _gen_strategic_recommendations(self, a: dict) -> str:
        """
        Generates the strategic recommendations section of the PDF report.

        This function prepares decision-making inputs such as dataset size,
        missing data, strong correlations, numeric averages, and outlier summaries.

        Args:
            a (dict): Analysis dictionary containing shape, missing summary, strong
            correlations, numeric means, and outlier statistics.

        Returns:
            str: Prioritized strategic recommendations based on verified analysis data.
        """
        grounded = {
            "shape": a["shape"],
            "missing_summary": a["missing_summary"],
            "strong_correlations": a["strong_correlations"][:5],
            "numeric_means": a["numeric_means"],
            "outlier_summary": [
                {"column": col, "iqr_outliers": s["iqr_outlier_count"], "pct": s["iqr_outlier_pct"]}
                for col, s in a["column_stats"].items() if s["iqr_outlier_count"] > 0
            ][:5],
        }
        return self.llm_narrative(
            "management consultant",
            "Provide 5-7 prioritised, actionable strategic recommendations backed by "
            "the exact figures provided. Format each as: [Priority] Recommendation — rationale.",
            grounded,
        )

    def _gen_key_questions(self, a: dict) -> str:
        """
        Generates sample questions for the PDF report.

        The questions are grounded in actual Excel table columns and dataset profile.
        They are question-only prompts intended for users to answer by inspecting,
        filtering, comparing, counting, summing, averaging, ranking, or calculating
        values from the Excel table.

        Returns:
            str: Dataset-specific sample questions with no answers.
        """
        grounded = {
            "columns": a["columns"], "numeric_means": a["numeric_means"], "strong_correlations": a["strong_correlations"][:5], "missing_summary": a["missing_summary"],
          
        }

        task = """
            Generate sample questions for this Excel table for my chatbot which answers of question based on provided excel table.

            Requirements:
            - Minimum 10 questions.
            - Decide the final number of questions based on dataset richness.
            - Questions only; no answers.
            - Number the questions.
            - Each question must reference actual column names where possible.
            - Do not include actual numerical answers in the questions.
            - Frame many questions so the answer requires a numerical lookup, calculation, comparison, count, sum, average, minimum, maximum, percentage, or ranking from the Excel table.
            - Use the dataset profile only to understand what question types are possible.
            - Do not expose computed statistics directly in the question.
            - Do not ask questions that cannot be answered from the Excel table.
            - Do not force question types unsupported by the data.
            - Avoid repeating the same question pattern.

            Include a balanced mix of applicable question types:
            1. Count-based questions
            2. Sum or total questions
            3. Average or mean questions
            4. Minimum and maximum questions
            5. Ranking questions
            6. Category-wise comparison questions
            7. Percentage or proportion questions
            8. Missing value lookup questions
            10. Outlier   lookup questions
            11. Correlation or relationship questions
            12. Time/date trend questions
            14. Filtered numerical lookup questions, where the user can calculate a numeric value after filtering by a category, region, product, customer, date, or segment

        """

        return self.llm_narrative(
            "expert Excel report assistant",
            task,
            grounded,
        )

    @staticmethod
    def _safe_paragraph(text: str, style) -> Paragraph:
        """
        Creates a ReportLab paragraph after escaping unsafe XML/HTML characters.

        ReportLab paragraphs interpret special characters such as ampersands and angle
        brackets. This helper escapes those characters so raw dataset text or LLM output
        does not break PDF rendering.

        Args:
            text (str): Text content to place inside the paragraph.
            style: ReportLab paragraph style used for rendering the text.

        Returns:
            Paragraph: Safe ReportLab Paragraph object ready to be added to the PDF.
        """
        safe = (
            text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
        )
        return Paragraph(safe, style)

    def _section_block(
        self,
        number: int,
        section_id: str,
        title: str,
        content: str,
        h1,
        body,
        mono,
    ) -> list:
        """
        Builds the PDF elements for one report section.

        The method adds a divider, section heading, formatted paragraphs, and spacing.
        Monospace styling is applied to technical sections such as dataset overview,
        missing values, and statistical summary.

        Args:
            number (int): Section number shown in the PDF.
            section_id (str): Internal section identifier used to choose formatting.
            title (str): Display title of the report section.
            content (str): Generated section text to render.
            h1: ReportLab style used for the section heading.
            body: ReportLab style used for normal paragraph text.
            mono: ReportLab style used for monospace technical text.

        Returns:
            list: ReportLab flowable elements representing the complete section block.
        """
        style = mono if section_id in self.MONO_SECTIONS else body
        elements = [
            HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc"), spaceAfter=4),
            Paragraph(f"{number}. {title}", h1),
        ]
        for para in content.split("\n"):
            para = para.strip()
            if para:
                elements.append(self._safe_paragraph(para, style))
        elements.append(Spacer(1, 0.3 * cm))
        return elements

    def build_pdf(
        self,
        output_path: str,
        per_table_data: List[Tuple[dict, dict, List[Tuple[str, str, str]]]],
    ) -> None:
        """
        Builds and writes the final PDF report.

        The PDF contains one top-level section per detected table.  Each section
        is headed with "Table N" and immediately followed by a three-row metadata
        block (sheet name, excel file name, table range).  Below the metadata
        block, ten numbered subsections present the LLM-generated analysis.

        Args:
            output_path (str): File path where the PDF should be written.
            per_table_data (List[Tuple[dict, dict, List[Tuple[str, str, str]]]]): 
                One entry per table, each a tuple of:
                  - meta dict  (sheet_name, excel_name, range, table_num)
                  - analysis dict  (Python-computed statistics for this table)
                  - sections_output list  (sid, title, content) × 10 subsections

        Returns:
            None
        """
        doc = SimpleDocTemplate(
            output_path, pagesize=A4,
            leftMargin=2 * cm, rightMargin=2 * cm,
            topMargin=2.5 * cm, bottomMargin=2 * cm,
            title="Data Analysis Report", author="Report Generator",
        )
        title_style, h1, body, mono = _build_styles()

        # Styles
        base = getSampleStyleSheet()
        sheet_heading_style = ParagraphStyle(
            "SheetHeading", parent=base["Heading1"],
            fontSize=14, leading=18,
            textColor=colors.HexColor("#1a1a2e"),
            spaceBefore=16, spaceAfter=4,
            fontName="Helvetica-Bold",
        )
        table_heading_style = ParagraphStyle(
            "TableHeading", parent=base["Heading2"],
            fontSize=12, leading=16,
            textColor=colors.HexColor("#16213e"),
            spaceBefore=10, spaceAfter=6,
            fontName="Helvetica-Bold",
        )
        subsection_heading_style = ParagraphStyle(
            "SubsectionHeading", parent=base["Heading3"],
            fontSize=11, leading=15,
            textColor=colors.HexColor("#16213e"),
            spaceBefore=10, spaceAfter=4,
            fontName="Helvetica-Bold",
        )

        story: list = []

        # Group tables by sheet so ## Sheet: X only prints once per sheet
        seen_sheets: set = set()

        # ── One section per table ─────────────────────────────────────
        for t_idx, (meta, analysis, sections_output) in enumerate(per_table_data, start=1):
            sheet_name = meta.get("sheet_name", "")
            table_num  = meta.get("table_num", t_idx)
            rng        = meta.get("range", "")

            # ── Sheet heading (printed once per unique sheet) ───────────────────
            if sheet_name not in seen_sheets:
                story.append(
                    HRFlowable(width="100%", thickness=1.5, color=colors.HexColor("#1a1a2e"), spaceAfter=4)
                )
                story.append(Paragraph(f"## Sheet: {sheet_name}", sheet_heading_style))
                seen_sheets.add(sheet_name)

            # ── Table heading with range inline ───────────────────────────────
            story.append(
                Paragraph(f"#### Table {table_num} (Range: {rng})", table_heading_style)
            )
            story.append(Spacer(1, 0.2 * cm))

            # ── 10 numbered subsections ───────────────────────────────────────────
            for sub_idx, (section_id, title, content) in enumerate(sections_output, start=1):
                style = mono if section_id in self.MONO_SECTIONS else body
                story.append(
                    HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#cccccc"), spaceAfter=3)
                )
                story.append(Paragraph(f"{t_idx}.{sub_idx} {title}", subsection_heading_style))
                for para in content.split("\n"):
                    para = para.strip()
                    if para:
                        story.append(self._safe_paragraph(para, style))
                story.append(Spacer(1, 0.25 * cm))

            story.append(PageBreak())

        doc.build(story)
        log.info("PDF written to '%s'", output_path)



if __name__ == "__main__":
    start = time.time()

    input_path = Path("Financial Sample-3.xlsx")
    output_path = input_path.with_suffix(".pdf")

    df = load_data(str(input_path))
    report = PDFGenerationReport(df)
    report.run(str(output_path))

    print(time.time() - start, "seconds")
