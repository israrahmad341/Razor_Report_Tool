# 💳 Razor Report Tool

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![MySQL](https://img.shields.io/badge/MySQL-4479A1?style=for-the-badge&logo=mysql&logoColor=white)
![LLMs](https://img.shields.io/badge/LLMs-FF6B35?style=for-the-badge&logo=openai&logoColor=white)

An automated reconciliation utility designed to aggregate and report on complex PG-2 and PG-5 payment gateway transaction streams.

## 🚀 Key Features

* **LLM-Powered Reconciliation**: Uses Prompt Engineering and Large Language Models to intelligently flag payment anomalies.
* **Automated Data Aggregation**: Seamlessly merges multiple transaction streams (PG-2 & PG-5) into a single unified view.
* **Smart Filtering & Counting**: Automatically counts active merchants and isolates unmatched/failed transactions.
* **ASM-Wise Export**: Exports fully formatted, professional Excel reports dynamically categorized by Area Sales Managers (ASM).

## 🛠️ Technology Stack

* **Language**: Python 3.x
* **Data Processing**: Pandas, OpenPyXL
* **Database**: MySQL, Amazon RDS
* **AI Integration**: Generative AI, Prompt Engineering

## 📝 Usage

1. Place your raw PG-2 and PG-5 transaction exports in the designated folder.
2. Ensure your `.env` is configured with database and API details.
3. Run `python Razor_Report_Tool.py` to generate the combined outputs.
