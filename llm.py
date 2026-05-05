import requests
import json
import re
from dotenv import load_dotenv
import os
load_dotenv()

TRITON_NEMOTRON_URL = os.getenv("TRITON_NEMOTRON_URL")
def call_llm(user_prompt, system_prompt):

    # Chat request
    chat_request = {
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt }
        ],
        "temperature": 0.1,
        "top_p": 0.9,
        "max_tokens": 32768,
        "chat_template_kwargs": {"enable_thinking":     True}
    }

    encoded = json.dumps(chat_request)

    # Build Triton inference request
    payload = {
        "inputs": [
            {
                "name": "chat_request",
                "shape": [1,1],
                "datatype": "BYTES",
                "data": [encoded]
            }
        ],
        "outputs": [
            {"name": "result"}
        ]
    }

    # Send request
    response = requests.post(TRITON_NEMOTRON_URL, json=payload)

    # Parse output
    result_json = response.json()
    output_text = result_json["outputs"][0]["data"][0]
    return output_text