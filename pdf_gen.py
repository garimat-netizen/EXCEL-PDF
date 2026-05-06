from __future__ import annotations

import re
import sys
import json
import logging
import tempfile
import textwrap
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
from preprocess_file import load_data, detect_numeric_columns, _build_styles

#-----PDF GENERATION MODULE -------#

class PDFGenerationReport:

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
        - Always include executive_summary and strategic_recommendations.
        - Return NO prose, NO markdown, NO explanation — only the JSON array.
    """)

    MONO_SECTIONS = {"dataset_overview", "missing_values", "statistical_summary"}

    def __init__(self, df: pd.DataFrame) -> None:
        self.df = df.drop(columns=["__sheet__", "__table__"], errors="ignore")

    def run(self, output: str = "report.pdf") -> None:
        df = self.df

        log.info("Running statistical analysis...")
        analysis = self.analyze_data(df)
        log.info(
            "Analysis complete: %d numeric cols, %d categorical cols, "
            "%d strong correlations, %d outlier cols",
            len(analysis["numeric_columns"]),
            len(analysis["categorical_columns"]),
            len(analysis["strong_correlations"]),
            sum(1 for s in analysis["column_stats"].values() if s["iqr_outlier_count"] > 0),
        )

        log.info("Asking LLM to select the 10 most relevant sections...")
        selected_ids = self.select_sections(analysis)
        log.info("LLM selected: %s", selected_ids)

        catalogue_map = {s["id"]: s for s in self.SECTION_CATALOGUE}

        with tempfile.TemporaryDirectory():
            log.info("Generating structured dataset context...")
            llm_context = self.generate_llm_context(df)

            log.info("Generating content for %d sections...", len(selected_ids))
            sections_output = []

            for sid in selected_ids:
                if sid not in catalogue_map:
                    log.warning("Unknown section id '%s', skipping.", sid)
                    continue

                entry = catalogue_map[sid]
                log.info("  Generating: %s", entry["title"])

                try:
                    fn = getattr(self, entry["fn_name"])
                    content = fn({**analysis, "llm_context": llm_context})
                except Exception:
                    log.warning("  Section '%s' failed:\n%s", sid, traceback.format_exc())
                    content = "[Section generation failed — see logs]"

                if not content or not content.strip():
                    content = "[No content generated for this section]"

                sections_output.append((sid, entry["title"], content))

            self.build_pdf(output, analysis, sections_output)

        print(f"\n  Report saved : {output}")
        print(f"  Sections     : {len(sections_output)}")
        for i, (sid, title, _) in enumerate(sections_output, 1):
            print(f"    {i:2}. {title}")

    def clean_llm_output(self, text: str) -> str:
        

        parts = text.split("</think>", 1)
        return parts[1].strip() if len(parts) > 1 else text


    def safe_llm(self, system: str, user: str, fallback: str = "[Content unavailable]") -> str:
        try:
            result = call_llm(system, user)

            result = self.clean_llm_output(result)

            return result.encode("latin-1", errors="replace").decode("latin-1")
        except Exception as exc:
            log.warning("LLM call failed: %s", exc)
            return fallback

    def llm_narrative(self, role: str, task: str, grounded_data: dict) -> str:
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
        try:
            v = float(x)
            return None if (np.isnan(v) or np.isinf(v)) else round(v, 6)
        except Exception:
            return None

    def analyze_data(self, df: pd.DataFrame) -> dict:
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
        known = {s["id"] for s in self.SECTION_CATALOGUE}
        match = re.search(r"\[.*?\]", raw, re.DOTALL)

        if not match:
            log.warning("Could not parse section IDs from LLM. Using defaults.")
            return defaults

        try:
            ids = json.loads(match.group())
            ids = list(dict.fromkeys(i for i in ids if i in known))[:10]
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
        rows = [
            f"{'Column':<30} {'Count':>7} {'Mean':>14} {'Median':>14} "
            f"{'Std':>14} {'Min':>14} {'Max':>14} {'Skew':>8}"
        ]
        rows.append("-" * 113)
        for col, s in a["column_stats"].items():
            def f(v):
                return f"{v:>14,.4f}" if v is not None else f"{'N/A':>14}"
            rows.append(
                f"{col:<30} {s['count']:>7,} {f(s['mean'])} {f(s['median'])} "
                f"{f(s['std'])} {f(s['min'])} {f(s['max'])} "
                f"{str(s['skewness'] or 'N/A'):>8}"
            )
        return "\n".join(rows)

    def _gen_trend_analysis(self, a: dict) -> str:
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
        if not a["categorical_stats"]:
            return "No categorical columns found in this dataset."
        return self.llm_narrative(
            "market analyst",
            "Analyse the categorical column distributions. Identify dominant categories, "
            "potential concentration risks, and segment-level insights.",
            {"categorical_stats": a["categorical_stats"]},
        )

    def _gen_kpi_analysis(self, a: dict) -> str:
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
        grounded = {
            "columns": a["columns"],
            "numeric_means": a["numeric_means"],
            "strong_correlations": a["strong_correlations"][:5],
            "missing_summary": a["missing_summary"],
        }
        task = """
Generate 10 specific questions based on actual columns and values.

Rules:
- Each question must reference a column name
- No generic business questions
- Keep each under 15 words

If insufficient data:
"Insufficient data for questions"
"""
        return self.llm_narrative(
            "senior analyst",
         task, 
            grounded,
        )

    @staticmethod
    def _safe_paragraph(text: str, style) -> Paragraph:
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
        analysis: dict,
        sections: List[Tuple[str, str, str]],
    ) -> None:
        doc = SimpleDocTemplate(
            output_path, pagesize=A4,
            leftMargin=2 * cm, rightMargin=2 * cm,
            topMargin=2.5 * cm, bottomMargin=2 * cm,
            title="Data Analysis Report", author="Report Generator",
        )
        title_style, h1, body, mono = _build_styles()
        story: list = []

        story.append(Spacer(1, 2 * cm))
        story.append(Paragraph("Data Analysis Report", title_style))
        story.append(Spacer(1, 0.5 * cm))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#1a1a2e")))
        story.append(Spacer(1, 0.5 * cm))

        cols_display = ", ".join(analysis["columns"])
        if len(cols_display) > 120:
            cols_display = cols_display[:120] + "..."

        meta_data = [
            ["Shape", f"{analysis['shape']['rows']:,} rows  x  {analysis['shape']['cols']} columns"],
            ["Columns", cols_display],
            ["Numeric", str(len(analysis["numeric_columns"])) + " columns"],
            ["Categorical", str(len(analysis["categorical_columns"])) + " columns"],
            ["Missing", f"{analysis['missing_summary']['missing_pct']} % of all cells"],
            ["Duplicates", f"{analysis['duplicate_rows']:,} rows"],
            ["Sections", f"{len(sections)} (LLM-selected from {len(self.SECTION_CATALOGUE)} candidates)"],
        ]
        meta_table = Table(meta_data, colWidths=[3.5 * cm, 13.5 * cm])
        meta_table.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#16213e")),
            ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#f0f0f0"), colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cccccc")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(meta_table)
        story.append(PageBreak())

        for idx, (section_id, title, content) in enumerate(sections, start=1):
            story.extend(self._section_block(idx, section_id, title, content, h1, body, mono))

        doc.build(story)
        log.info("PDF written to '%s'", output_path)


import time

if __name__ == "__main__":
    start = time.time()
    INPUT_FILE = "Financial Sample-3.xlsx"
    OUTPUT_FILE = Path(INPUT_FILE).stem + ".pdf"

    df = load_data(INPUT_FILE)
    report = PDFGenerationReport(df)
    report.run(OUTPUT_FILE)

    print(time.time() - start, "seconds")
