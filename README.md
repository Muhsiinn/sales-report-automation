# 📊 Sales Report Automation

This project is an initial version of an automated reporting and email delivery system. It processes sales data from the [Online Retail dataset](https://archive.ics.uci.edu/dataset/352/online+retail) to generate insightful Excel reports and email them as HTML summaries.

---

## 🚀 Features

- 🏗️ Generates multi-sheet Excel reports with:
  - Revenue summaries
  - Top-selling items
  - Most cancelled items and customers
- 📧 Sends professional HTML email summaries with embedded tables
- 🔒 Secure `.env` setup for handling email credentials
- 📝 Modular and extensible Python scripts

---

## 📂 Project Structure

sales-report-automation/
├── data/ # Raw and processed data
├── reports/ # Generated Excel reports
├── scripts/
│ ├── generate_report.py # Report creation logic
│ └── send_email.py # Email automation script
├── notebooks/ # Jupyter workbench for exploration
├── .env # Email credentials (ignored in Git)
├── .gitignore
└── README.md


---

## 🧪 Data Source

The dataset is sourced from:
[Online Retail Data Set - UCI Machine Learning Repository](https://archive.ics.uci.edu/dataset/352/online+retail)

---

## 💻 How to Run

1️⃣ Clone the repository:
```bash
git clone https://github.com/yourusername/sales-report-automation.git
cd sales-report-automation

2️⃣ Install dependencies:
pip install -r requirements.txt

3️⃣ Create a .env file in the root directory:
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASSWORD=your_app_password

4️⃣ Generate the report:
python scripts/generate_report.py

5️⃣ Send the report via email:
python scripts/send_email.py


🧑‍💻 Jupyter Workbench
For a deeper understanding of the data flow and to test custom HTML email templates, check the notebooks/ directory. The Jupyter notebook includes:

Explorations and tests of the report generation process

A more aesthetic version of the HTML email summary, with better table styling and layout


🔮 Future Plans
📅 Add dynamic time-window analysis (e.g., last 6 months, weekly trends)

📈 Generate natural language summaries of sales trends

📬 Integrate scheduled reporting with cron/Task Scheduler

🌐 Support multiple data sources and real-time pipelines

🎨 Enhance HTML email design and add embedded charts


🏗️ Current Status
This is a proof-of-concept project created to explore automated reporting and email delivery using Python. It lays the groundwork for future expansions and more robust business intelligence pipelines.

