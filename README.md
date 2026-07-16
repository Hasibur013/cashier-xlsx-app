---
title: Chayabithi Cafe SMS to Excel
emoji: ☕
colorFrom: blue
colorTo: green
sdk: gradio
sdk_version: 4.44.0
app_file: app.py
pinned: false
---

# ☕ Chayabithi Cafe — Cashier SMS → Excel Updater

The cashier of Chayabithi Cafe (Rangpur) sends a daily Bengali sales message.
This app reads that message, fills the matching **date column** of the monthly
Excel workbook, and gives back the updated file to download.

## How it works
1. Paste the cashier's SMS.
2. Upload the monthly `.xlsx` workbook.
3. Pick a parsing engine and press **প্রসেস করুন (Process)**.
4. Download the updated workbook. Use **ব্যাকএন্ড ক্লিয়ার** to wipe generated files.

The message date (e.g. `১৫/০৭/২৬`) selects the column; each item's quantity is
written to its `Item Quentity` row, and the sheet's own `Total taka` formulas do
the maths. The kids ride **Jumping** is written to its special rows.

## Parsing engines
- **Offline (default):** a deterministic regex parser (`parser_core.py`). Fast,
  free, no network, and safest for bookkeeping numbers.
- **AI (Groq LLM):** optional. Set `GROQ_API_KEY`. Falls back to the offline
  parser automatically if the key is missing or the call fails.

> ⚠️ **Model note:** `llama-3.1-70b-versatile` is **decommissioned** on Groq and
> `llama-3.3-70b-versatile` is scheduled for shutdown on **2026-08-16**. The app
> defaults to `openai/gpt-oss-120b`. Override with the `GROQ_MODEL` env var.

## Environment variables
| Variable | Purpose | Default |
| --- | --- | --- |
| `GROQ_API_KEY` | Groq key (AI engine only) | _(empty)_ |
| `GROQ_MODEL` | Groq model id | `openai/gpt-oss-120b` |

## Run locally
```bash
pip install -r requirements.txt
python app.py
```

## Deploy free on Hugging Face Spaces
1. Create a new **Gradio** Space at <https://huggingface.co/new-space>.
2. Push this repo to the Space (see below).
3. In **Settings → Variables and secrets**, add `GROQ_API_KEY` (and optionally
   `GROQ_MODEL`) — only if you want the AI engine.
4. The Space builds from `requirements.txt` and launches `app.py` automatically.

> 🔒 **Privacy:** the real workbook is git-ignored by default because a public
> Space is world-readable. Users upload the file at runtime, so it never needs
> to live in the repo. If you want a shareable template, remove the workbook
> line from `.gitignore`, or make the Space **private**.
