# ZimGoodies — Business Tracker

ZimGoodies is a small retail business run by three equal partners (33.3% share each). Stock is sourced from Zimbabwe (costs in USD) and sold in Tanzania (revenue in Tanzanian Shillings, TZS). This Google Sheets + Apps Script tool tracks daily sales, stock, debts, expenses, and partner profit splits — all from a shared spreadsheet accessible on any phone or computer. No technical knowledge required: just open the sheet, fill in data, and press a button to get a WhatsApp-ready daily report.

---

## Prerequisites

- A Google account (free — gmail.com works fine)
- No installations, no downloads, nothing to pay

---

## Setup (step by step)

1. Go to [sheets.google.com](https://sheets.google.com) and click the **+** button to create a new blank spreadsheet.

2. Click the title at the top (it says "Untitled spreadsheet") and rename it **ZimGoodies**.

3. Click **Extensions** in the menu bar, then click **Apps Script**.

4. A new tab opens showing a code editor. You will see some default code (usually `function myFunction() {}`). **Select all of it and delete it** — the editor should be completely empty.

5. Open the file `Code.gs` from this repository. Copy **the entire contents** of that file.

6. Click inside the empty Apps Script editor and paste the code.

7. Press **Ctrl+S** (Windows) or **Cmd+S** (Mac) to save. Give the project any name when prompted (e.g. "ZimGoodies").

8. In the top toolbar of the Apps Script editor, find the dropdown that says **"Select function"** and choose **`setupSpreadsheet`** from the list.

9. Click the **▶ Run** button.

10. A popup will appear asking you to **Review Permissions**. Click it, choose your Google account, click **Advanced**, then **Go to ZimGoodies (unsafe)** (this is normal for personal scripts — it just means it hasn't been through Google's app review), and finally click **Allow**.

11. Go back to your spreadsheet tab. After a few seconds, all 8 sheets will be created and a green message will appear at the bottom saying **"✅ ZimGoodies is ready!"**

12. Click **Share** (top right of the spreadsheet) and add all partners' email addresses as **Editors** so everyone can use it.

13. **Bookmark** the spreadsheet URL in your phone browser so you can open it quickly each day.

---

## How to update the exchange rate

1. Open the spreadsheet and click the **SETTINGS** tab.
2. Find the row that says **💱 USD → TZS Exchange Rate**.
3. Click the yellow cell in that row (column C) and type the new rate — for example `3500`.
4. Press Enter. All calculations in the whole spreadsheet update automatically.

---

## How to add a new product

1. Open the **SETTINGS** tab.
2. Scroll down to the **📦 Product List** table.
3. Add a new row at the bottom of the table: type the product name, cost in USD, selling price in TZS, and the low stock threshold quantity.
4. Then go to the **STOCK** tab and add a new row for the product manually (copy the format of an existing row).
5. To update the dropdowns across all sheets, click **🛒 ZimGoodies** in the menu bar → **⚙️ Setup / Reset Sheets** → **Yes**. Existing data is safe — setup only creates things that are missing.

---

## How to record a sale

1. Open the **SALES LOG** tab.
2. Click the first empty row below the last entry.
3. Fill in:
   - **Date** — type today's date (format: DD/MM/YYYY)
   - **Customer Name** — or leave blank for walk-in customers
   - **Product** — click the cell and pick from the dropdown
   - **Qty** — how many units sold
   - **Unit Price TZS** — this auto-fills, but you can change it if you gave a discount
   - **Total TZS** — auto-calculated (Qty × Unit Price)
   - **Amount Paid TZS** — how much the customer paid right now
   - **Balance Owed TZS** — auto-calculated (if positive, the row turns red)
   - **Recorded By** — pick your name from the dropdown
   - **Notes** — optional, e.g. "paid half now, rest Thursday"
4. If a customer owes money (balance > 0), also add them to the **DEBTS** tab for tracking.

---

## How to record new stock arriving

When stock comes in from Zimbabwe:

1. Open the **STOCK RECEIVED** tab and add a row:
   - Date, Product (dropdown), Qty Received, Cost Per Unit USD, Supplier/Notes, your name
   - Total Cost USD fills automatically.

2. Open the **STOCK** tab and find the same product row.
   - Update **Current Qty** (add the new quantity to whatever was there).
   - Update **Last Updated** to today's date.

That's it. The Dashboard will pick up the new cost automatically when you refresh it.

---

## How to get the daily WhatsApp report

1. Open the spreadsheet.
2. Click **🛒 ZimGoodies** in the menu bar at the top.
3. Click **📊 Daily WhatsApp Report**.
4. A box pops up with the full report text already formatted.
5. Click **📋 Copy to Clipboard** (or click inside the box, press Ctrl+A, then Ctrl+C).
6. Open WhatsApp and paste into your partners group chat.

---

## How to refresh the Dashboard

Click **🛒 ZimGoodies** in the menu → **🔄 Refresh Dashboard**. The DASHBOARD tab updates with all current totals, profit shares, stock alerts, and outstanding debts.

---

## Sheet summary

| Sheet | Purpose |
|---|---|
| SETTINGS | Edit partners, exchange rate, products, categories |
| DASHBOARD | Live summary — refresh via menu |
| SALES LOG | Daily sales entry |
| STOCK | Current inventory levels |
| DEBTS | Customers who owe money |
| EXPENSES | All running costs (TZS or USD) |
| PARTNER ACCOUNT | Partner withdrawals and contributions |
| STOCK RECEIVED | Purchase log when new stock arrives |

---

## GitHub

[https://github.com/isabelmoyo/zim_goodies](https://github.com/isabelmoyo/zim_goodies)
