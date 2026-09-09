# Excel LLM Bridge

Ask a local LLM (via [Ollama](https://ollama.com)) a question directly from an Excel cell, and get the answer back in another cell — no add-ins, no plugins, just VBA + a small Python script + two text files.

## How it works

```
 ┌────────────┐   writes    ┌───────────────┐   reads     ┌─────────────────┐
 │  Excel      │──────────▶│ question.txt  │────────────▶│  Python watcher  │
 │  (VBA       │            └───────────────┘              │  (watcher.py)   │
 │   macro)    │                                             └────────┬────────┘
 │             │            ┌───────────────┐   writes               │
 │             │◀──────────│  answer.txt   │◀───────────────────────┘
 └────────────┘   reads    └───────────────┘        POST /api/generate
                                                              │
                                                              ▼
                                                     ┌─────────────────┐
                                                     │  Ollama server   │
                                                     │  (e.g. qwen3:14b)│
                                                     └─────────────────┘
```

1. You type a question into cell **I17**.
2. The `AutoRefresh` VBA macro writes it to `question.txt`.
3. `watcher.py` notices the new question, sends it to your Ollama server, and writes the reply to `answer.txt`.
4. The VBA macro reads `answer.txt` and drops the result into cell **I18**.

The whole loop runs every ~2-3 seconds, so it feels close to a live chat inside a spreadsheet cell.

## Project structure

```
excel-llm-bridge/
├── python/
│   └── watcher.py        # Polls question.txt, calls Ollama, writes answer.txt
├── vba/
│   └── AutoRefresh.bas   # Watches I17, writes question.txt, reads answer.txt into I18
└── README.md
```

## Requirements

- Windows + Microsoft Excel (the macro uses VBA `Open`/`Print`/`Dir` file I/O)
- Python 3.8+
- The [`requests`](https://pypi.org/project/requests/) package: `pip install requests`
- A running [Ollama](https://ollama.com) server with a model pulled, e.g.:
  ```
  ollama pull qwen3:14b
  ollama serve
  ```

## Setup

1. **Clone the repo**
   ```
   git clone https://github.com/Ilia-Tsituk/excel-llm-bridge.git
   cd excel-llm-bridge
   ```

2. **Configure the Python watcher.**
   Open `python/watcher.py` and set (or export as environment variables):
   - `OLLAMA_SERVER` – URL of your Ollama server (default `http://127.0.0.1:11434`)
   - `OLLAMA_MODEL` – model name to use (default `qwen3:14b`)
   - `BRIDGE_FOLDER` – folder where `question.txt` / `answer.txt` will live (default `C:\Users\hello\Desktop`)

3. **Import the macro into Excel.**
   - Open your workbook, press `Alt+F11` to open the VBA editor.
   - Import `vba/AutoRefresh.bas` (File → Import File...), or copy-paste its contents into a new module.
   - In the macro, make sure `filePath` matches the `BRIDGE_FOLDER` you set in step 2.

4. **Run it.**
   - Start the Python watcher: `python python/watcher.py`
   - In Excel, run the `AutoRefresh` macro once (e.g. from the Developer tab, or bind it to a button / `Workbook_Open`).
   - Type a question in cell `I17` — the answer will appear in `I18` a few seconds later.
   - Run `StopRefresh` to stop the polling loop.

## Notes & limitations

- Communication happens through two plain-text files, so it's simple but not instantaneous — expect a couple of seconds of latency per exchange.
- Only one question/answer pair is tracked at a time; a new question overwrites the previous exchange.
- The prompt template in `watcher.py` is tuned for short, calculation-style answers (handy for spreadsheet cells) — edit `PROMPT_TEMPLATE` to change the style, verbosity, or language of responses.
- Files are read/written using the Windows "ANSI" codepage to match Excel's default VBA file I/O; if you use non-ASCII characters, make sure the codepage matches on both ends.
- This is a lightweight personal/demo project rather than a hardened integration — there's no authentication or input sanitization, so only point it at a trusted local Ollama instance.

## License

This project is released under the [MIT License](LICENSE).
