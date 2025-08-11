# PPTX Inconsistency Detector

A terminal Python tool that analyzes multi-slide PowerPoint (`.pptx`) decks for factual or logical inconsistencies — including conflicting numbers, contradictory textual claims, and timeline mismatches. The tool extracts text from slides (and OCRs embedded images), sends the consolidated content to **Gemini 2.5 Flash** for semantic analysis, and outputs a structured JSON report.

---

## Features & Functionality
- Full-deck analysis using an LLM for:
  - Conflicting numerical data
  - Contradictory textual claims
  - Timeline mismatches
- Slide text extraction with `python-pptx`.
- Image OCR for text embedded in pictures using Tesseract OCR.
- Context-aware detection — avoids false positives like “compressed from 1.5 MB to 45 KB”.
- Structured JSON output with slide numbers, issue type, description, reason, and confidence score.
- Readable terminal summary alongside machine-readable report.

---

## Installation
```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
# venv\Scripts\activate         # Windows PowerShell

pip3 install -r requirements.txt
```

---

## Environment Variables
```bash
export GEMINI_API_KEY="YOUR_KEY_HERE"
export GEMINI_ENDPOINT="https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
```
Get a free API key here: [Google AI Studio](https://aistudio.google.com/app/apikey)

---

## Usage
```bash
python3 enhanced_pptx_detector.py -i NoogatAssignment.pptx -o report.json
```
Options:
- `-i`, `--input` — path to `.pptx` file (required)
- `-o`, `--out` — path to JSON report output (default: `report.json`)
- `--no-ocr` — skip OCR step

---

## Example Output (JSON)
```json
[
  {
    "type": "textual_contradiction",
    "slide_a": 2,
    "slide_b": 5,
    "snippet_a": "Market is highly competitive, with many players.",
    "snippet_b": "We have few competitors and large market share.",
    "confidence": 0.87,
    "explanation": "Contradiction: one slide says many players, the other says few competitors."
  }
]
```

---

## How It Works
1. Extract text from all `.pptx` slides.
2. Extract images and run OCR to recover text.
3. Send all content to Gemini 2.5 Flash for analysis.
4. Parse LLM output into a JSON report and print a summary.

---

## Scalability, Generalisability & Robustness
- Can be extended for large decks with chunking and parallel OCR.
- Works with most `.pptx` decks regardless of topic.
- LLM reasoning reduces false positives compared to regex-only methods.

---

## Limitations
- Only `.pptx` input supported — no direct image folder or PDF input.
- Extracts images from `.pptx` and runs OCR, but cannot take standalone image files.
- No automatic chunking for large decks; may hit Gemini’s context limit.
- OCR accuracy depends on Tesseract quality.
- No tests or CI pipeline included.
- Limited metric normalization.
- No Dockerfile or packaged release.
- Sends slide text to an external API — handle sensitive data with care.

---

## Future Improvements
- Direct image/PDF input.
- Automatic chunking for large decks.
- Better table/chart text extraction.
- CI/CD with unit tests.
- Dockerized build.
- Offline rule-based mode.
