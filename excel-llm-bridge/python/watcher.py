"""
watcher.py

Watches a plain-text "question" file for changes, forwards the question to a
local Ollama server, and writes the model's answer back to a plain-text
"answer" file. The companion Excel VBA macro (see vba/AutoRefresh.bas)
writes the question file and reads the answer file, which gives you a
simple "ask an LLM from a spreadsheet cell" workflow with no plugins or
add-ins required.

Requirements:
    pip install requests

Usage:
    1. Start your Ollama server (or point SERVER at a remote one).
    2. Set SERVER, MODEL and FOLDER below (or via environment variables).
    3. Run this script: python watcher.py
    4. Open the workbook and run the AutoRefresh macro.
"""

import time
import requests
import os
import json
import re

# --- Configuration ---------------------------------------------------------
# You can hardcode these values or override them with environment variables.
SERVER = os.environ.get("OLLAMA_SERVER", "http://127.0.0.1:11434")
MODEL = os.environ.get("OLLAMA_MODEL", "qwen3:14b")
FOLDER = os.environ.get("BRIDGE_FOLDER", r"C:\Users\hello\Desktop")

QUESTION_FILE = os.path.join(FOLDER, "question.txt")
ANSWER_FILE = os.path.join(FOLDER, "answer.txt")

# Prompt template sent to the model. Edit this to change the model's
# response language/style (kept short so it fits nicely back into a cell).
PROMPT_TEMPLATE = (
    "Answer in English. Give ONLY the calculation and the final answer "
    "in 1-3 lines maximum. No explanations, no extra text:\n\n{question}"
)

POLL_INTERVAL_SECONDS = 2
# ----------------------------------------------------------------------------

print("Watcher started. Waiting for a question...")

last_question = None

while True:
    try:
        if os.path.exists(QUESTION_FILE):
            with open(QUESTION_FILE, "r", encoding="ansi") as f:
                question = f.read().strip()

            if question and question != last_question:
                print(f"New question: {question}")
                print("Sending request to the server...")

                r = requests.post(
                    f"{SERVER}/api/generate",
                    json={
                        "model": MODEL,
                        "prompt": PROMPT_TEMPLATE.format(question=question),
                        "stream": False,
                        "think": False,
                    },
                    timeout=300,
                )

                text = r.text.strip()
                result = None
                for line in text.splitlines():
                    try:
                        obj = json.loads(line)
                        if "response" in obj and obj["response"]:
                            result = obj["response"]
                            break
                    except json.JSONDecodeError:
                        continue

                if not result:
                    result = json.loads(text)["response"]

                # Strip <think>...</think> blocks in case the model returns them
                result = re.sub(r"<think>.*?</think>", "", result, flags=re.DOTALL).strip()

                print(f"Answer received: {result[:80]}...")

                with open(ANSWER_FILE, "w", encoding="ansi") as f:
                    f.write(result)

                last_question = question
                print("Done - the answer will appear in Excel within a few seconds")

    except Exception as e:
        print("Error:", type(e).__name__, e)

    time.sleep(POLL_INTERVAL_SECONDS)
