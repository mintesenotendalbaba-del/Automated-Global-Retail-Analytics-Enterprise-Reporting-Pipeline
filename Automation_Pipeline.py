"""
Retail Intelligence Automation Suite
------------------------------------
Automates Excel dashboard updates and report delivery:
1. Loads and cleans CSV data
2. Merges datasets into a master report
3. Refreshes Excel dashboard
4. Exports dashboard as PDF
5. Emails report automatically via Gmail
"""

# === Imports ===
import pandas as pd
import glob, os
import xlwings as xw
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# === Step 1: Point to your folder ===
folder_path = r"C:\Users\Minte\Documents\python automation\excel automation\PROJECT 02"

# === Step 2: Find all CSV files ===
files = glob.glob(os.path.join(folder_path, "*.csv"))
print("Files found:", files)

# === Step 3: Load each file into Pandas ===
data = {}
for f in files:
    name = os.path.basename(f).replace(".csv", "")
    data[name] = pd.read_csv(f, encoding="ISO-8859-1")
    print(f"Loaded {name} with {len(data[name])} rows")

# Quick test
print(data["Sales"].head())

# === Step 4: Cleaning & Transformation ===
print("\n--- Step 2: Cleaning & Mixing ---")

sales = data["Sales"]
products = data["Products"]
customers = data["Customers"]
stores = data["Stores"]
rates = data["Exchange_Rates"]

# Clean column names
for df in [sales, products, customers, stores, rates]:
    df.columns = df.columns.str.strip().str.lower()

# Drop duplicates
customers = customers.drop_duplicates(subset="customerkey")
products = products.drop_duplicates(subset="productkey")
stores = stores.drop_duplicates(subset="storekey")
rates = rates.drop_duplicates(subset="currency")

# Clean numeric columns
for col in ["unit cost usd", "unit price usd"]:
    if col in products.columns:
        products[col] = (
            products[col]
            .astype(str)
            .str.replace(r"[\$,]", "", regex=True)
            .astype(float)
        )

# ✅ Add $10 to Unit Price
if "unit price usd" in products.columns:
    products["unit price usd"] = products["unit price usd"] + 10

# Merge datasets
merged = (
    sales[["order number","customerkey","storekey","productkey","quantity","currency code"]]
    .merge(products[["productkey","product name","unit cost usd","unit price usd"]], on="productkey", how="left")
    .merge(customers[["customerkey","name","country"]], on="customerkey", how="left")
    .merge(stores[["storekey","state","country"]], on="storekey", how="left")
    .merge(rates[["currency","exchange"]], left_on="currency code", right_on="currency", how="left")
)

# Calculate metrics
merged["amount_usd"] = merged["quantity"] * merged["unit price usd"] * merged["exchange"]
merged["profit_margin"] = merged["amount_usd"] - (merged["unit cost usd"] * merged["quantity"])

# Save master dataset
master_path = os.path.join(folder_path, "Master_Sales_Report.csv")
if os.path.exists(master_path):
    os.remove(master_path)
    print("Old Master_Sales_Report.csv deleted.")
merged.to_csv(master_path, index=False)
print("✅ New Master dataset created successfully!")

# === Step 5: Delete old raw CSVs ===
for file in glob.glob(os.path.join(folder_path, "*.csv")):
    if "Master_Sales_Report.csv" not in file:
        os.remove(file)
        print(f"Deleted old file: {file}")

# === Step 6: Refresh Excel Dashboard ===
dashboard_path = os.path.join(folder_path, "Dashboard_Final.xlsm")
pdf_path = os.path.join(folder_path, "Dashboard_Report.pdf")

wb = xw.Book(dashboard_path)
wb.api.RefreshAll()
wb.save()

# Delete old PDF
if os.path.exists(pdf_path):
    os.remove(pdf_path)
    print("🗑️ Old dashboard PDF deleted.")

# Export dashboard sheet as PDF
wb.sheets["Dashboard"].api.ExportAsFixedFormat(0, pdf_path)
wb.close()
print("📊 Dashboard refreshed and exported as PDF.")

# === Step 7: Email Report via Gmail ===
sender = "mintesenotendalbaba@gmail.com"
password = "ceaf vcxk fxbf syww"   # Gmail app password
receiver = "mentefx1@gmail.com"

msg = MIMEMultipart()
msg['From'] = sender
msg['To'] = receiver
msg['Subject'] = "Updated Sales Dashboard"
msg.attach(MIMEText("Hello Mintesenot,\n\nPlease find attached the latest dashboard report.", 'plain'))

with open(pdf_path, "rb") as attachment:
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(pdf_path)}')
    msg.attach(part)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender, password)
server.send_message(msg)
server.quit()

print("📧 Dashboard emailed successfully via Gmail")
