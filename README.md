# Automated-Global-Retail-Analytics-Enterprise-Reporting-Pipeline

## 📊 What This Project Does
In many businesses, important numbers are trapped in different, separate files (Sales, Products, Stores, Customers, and Exchange Rates). Because this retailer sells globally, a person used to spend hours every single month manually copying data, fixing formatting errors, converting foreign currencies, and building charts by hand just to see their net profit.

I built a Python automation system that acts like a hands-free conveyor belt. It takes all those separate files, cleans them, glues them together, runs the business math, updates an Excel dashboard, and emails a finished PDF report to management—all in under 5 seconds.

---

## 🚀 How It Works (The Pipeline)

1. The Data Glue: Python automatically opens all 5 separate CSV files and matches them up perfectly using ID numbers (so it knows exactly which product was sold in which store location).
2. Global Currency Conversion: The code checks the Exchange_Rates file and automatically converts international sales into standard US Dollars (USD) so the math is accurate.
3. The Math & Logic: It automatically calculates Total Revenue, Cost of Goods Sold (COGS), and Net Profit Margin for every single transaction.
4. Hands-Free Reporting: Using xlwings, the script pushes the clean data into Excel in the background, updates the charts, saves the page as a professional PDF, and emails it straight to the boss's inbox.

---

## 🛠️ Tech Used
* Python (The main engine)
* Pandas (For cleaning and matching the data files)
* Xlwings (For talking to Excel, creating charts, and making PDFs)
* SMTP / Email Protocols (For sending the automated emails)

---

## 📂 Files Included
* Automation_Pipeline.py — The core Python automation code.
* Sales.csv, Products.csv, Stores.csv, Customers.csv, Exchange_Rates.csv — The raw, messy data files used for the project.
* README.md — This project explanation.

---

## 💻 How to Run It
1. Put your raw CSV data files into the same folder as the Python script.
2. Run the code:
   `bash
   python Automation_Pipeline.py

