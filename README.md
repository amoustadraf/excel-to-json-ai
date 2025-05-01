# Excel to JSON for AI

Convert Excel files into structured JSON that includes cell values, formatting (bold, italic, colors, borders), and sheet structure.

## Purpose

This tool makes Excel files easier for AI agents to understand and interact with.  
It is designed to support macro generation, workflow automation, or any context where an AI needs to reason about Excel content.

## What It Does

- Reads `.xlsx` or `.xlsm` files
- Extracts:
  - Cell values
  - Font styles (bold, italic, underline, strikethrough)
  - Font and fill colors
  - Border styles
- Outputs a clean JSON file suitable for AI input
- Optionally recreates an Excel file from the JSON (to verify output visually)

## Example Use Case

You upload an Excel file, the tool generates `output.json`, and an AI (ChatGPT, Gemini, Anthropic, etc.) can then use that structured data to generate actions or code.  
The actions can be tested or previewed using the `sandbox.py` script to see the file created from the JSON.

## How to Run

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
1. Clone the repository
git clone https://github.com/your-username/excel-to-json-ai
cd excel-to-json-ai

2. Install dependencies
pip install -r requirements.txt

3. Run the scanner script to generate output.json
python script.py

4. (Optional) Recreate Excel file from output.json
python sandbox.py
