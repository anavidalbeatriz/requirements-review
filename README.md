# 🧠 AI Document Evaluator

This project automatically **evaluates `.docx` documents** using **OpenAI models** based on a set of customizable rules.  
It reads evaluation criteria from a `rules.json` file, analyzes each document, and exports the results to an Excel file.

---

## 🚀 Features

- 📄 Reads and processes `.docx` files.  
- ⚙️ Evaluates documents according to structured rules defined in `rules.json`.  
- 🤖 Uses the OpenAI API (`gpt-4o-mini` by default) for scoring and justifications.  
- 📊 Exports a summary of scores and justifications to an Excel spreadsheet.  
- 💡 Easy to adapt for academic, business, or compliance evaluations.

---

## 📂 Project Structure

├── documents/ # Folder containing the .docx files to evaluate
├── rules.json # JSON file with evaluation rules and criteria
├── .env # Contains your OpenAI API key
├── evaluation.xlsx # Output file with results
├── main.py # Main script
└── README.md


---

## ⚙️ Installation

### 1. Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/ai-document-evaluator.git
cd ai-document-evaluator

python -m venv venv
source venv/bin/activate    # macOS/Linux
venv\Scripts\activate       # Windows

pip install -r requirements.txt

OPENAI_API_KEY=your_api_key_here

Developed by: Ana Beatriz Vidal
License: MIT
Contact: anavidalbeatriz@gmail.com