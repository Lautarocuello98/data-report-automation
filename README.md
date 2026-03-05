# 📊 Data Report Automation

> A Python data processing and reporting tool that converts raw CSV/Excel files into cleaned datasets and professional Excel reports.

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests: pytest](https://img.shields.io/badge/tests-pytest-green.svg)](#-running-tests)

---

## ✨ Overview

Data Report Automation is a **Python automation tool** designed to transform raw business data into structured reports.

The system loads CSV or Excel files, cleans and normalizes the data, calculates business KPIs, and generates professional Excel reports automatically.

This project demonstrates:

- Data processing with pandas
- File automation
- Report generation with Excel
- CLI tool design
- Logging and structured workflows
- Automated testing with pytest

---

## 🚀 Features

| Feature | Description |
|-------|-------------|
| Multi-file processing | Process a single file or entire folder |
| Data cleaning | Handle duplicates, missing values, and formatting |
| KPI calculations | Revenue, cost, profit, and performance metrics |
| Excel reports | Generate formatted Excel reports |
| Charts | Automatically create visual summaries |
| Logging | Processing logs for traceability |
| Tests | Automated tests using pytest |

---

## 📦 Installation

### Requirements

- Python 3.10+
- pip

### Clone the repository

```bash
git clone https://github.com/Lautarocuello98/data-report-automation.git
cd data-report-automation
```

### Install dependencies

```bash
pip install -r requirements.txt
```

---

## 🎯 Quick Start

Run the tool with a folder containing CSV or Excel files:

```bash
python cli.py --input data/ --output reports/
```

Example output:

```
reports/
├── sales_report.xlsx
├── processing.log
└── charts/
```

---

## 📁 Project Structure

```
data-report-automation/
│
├── data/                # Input data files
├── reports/             # Generated reports
│
├── src/
│   ├── loader.py
│   ├── cleaner.py
│   ├── processor.py
│   ├── report_generator.py
│   └── charts.py
│
├── tests/
│
├── cli.py
├── config.json
├── requirements.txt
└── README.md
```

---

## ⚙️ Configuration

The system can be customized using `config.json`.

Options include:

- Column name mapping
- Cleaning rules
- Supported input formats
- Report settings

---

## 📊 Example Workflow

1️⃣ Load data files (CSV / Excel)  
2️⃣ Clean and normalize the dataset  
3️⃣ Compute KPIs such as revenue and profit  
4️⃣ Generate a structured Excel report  
5️⃣ Save logs and charts

---

## 🧪 Running Tests

Run all tests:

```bash
pytest -v
```

---

## 🧰 Technologies Used

- Python
- pandas
- openpyxl
- matplotlib
- pytest

---

## 📄 License

This project is licensed under the MIT License.

See the **LICENSE** file for details.

---

## 👨‍💻 Author

Lautaro Cuello

GitHub:  
https://github.com/Lautarocuello98

---

⭐ If you found this project useful, consider giving this repository a star.