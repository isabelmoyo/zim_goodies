# ZimGoodies — User Guide

For Isabel, Fra, and Rumbi.
This guide covers everything you need to run the business tracker day-to-day.
No technical knowledge needed.

---

## Table of Contents

1. [Opening the spreadsheet](#1-opening-the-spreadsheet)
2. [Overview of all sheets](#2-overview-of-all-sheets)
3. [Daily workflow — quick version](#3-daily-workflow--quick-version)
4. [Recording a sale](#4-recording-a-sale)
5. [Recording a debt (customer owes money)](#5-recording-a-debt-customer-owes-money)
6. [Marking a debt as paid](#6-marking-a-debt-as-paid)
7. [Recording an expense](#7-recording-an-expense)
8. [Logging new stock arriving from Zimbabwe](#8-logging-new-stock-arriving-from-zimbabwe)
9. [Updating stock quantities](#9-updating-stock-quantities)
10. [Recording a partner withdrawal or contribution](#10-recording-a-partner-withdrawal-or-contribution)
11. [Refreshing the Dashboard](#11-refreshing-the-dashboard)
12. [Getting the daily WhatsApp report](#12-getting-the-daily-whatsapp-report)
13. [Updating the exchange rate](#13-updating-the-exchange-rate)
14. [Changing product prices](#14-changing-product-prices)
15. [Adding a new product](#15-adding-a-new-product)
16. [Common questions](#16-common-questions)

---

## 1. Opening the spreadsheet

- Open your phone or computer browser and go to **[sheets.google.com](https://sheets.google.com)**
- Sign in with your Google account
- Find **ZimGoodies** in your recent files and tap/click it
- **Tip:** Bookmark it or add it to your home screen so you can open it in one tap

---

## 2. Overview of all sheets

At the bottom of the spreadsheet you will see coloured tabs. Here is what each one is for:

| Tab | Colour | What it's for |
|---|---|---|
| **SETTINGS** | Grey | Control panel — edit partners, prices, exchange rate |
| **DASHBOARD** | Blue | Live business summary — refresh to update |
| **SALES LOG** | Green | Record every sale here, every day |
| **STOCK** | Orange | Current stock levels — update qty when stock changes |
| **DEBTS** | Red | Customers who owe you money |
| **EXPENSES** | Yellow | All business costs (transport, airtime, packaging, etc.) |
| **PARTNER ACCOUNT** | Purple | Partner withdrawals and contributions |
| **STOCK RECEIVED** | Teal | Log when new stock arrives from Zimbabwe |

---

## 3. Daily workflow — quick version

Every day you will typically do these things:

1. **Record each sale** → SALES LOG
2. **Record any expenses** → EXPENSES
3. **Update debts** if a customer pays what they owed → DEBTS
4. At end of day: **Refresh Dashboard** → menu → 🔄 Refresh Dashboard
5. **Send WhatsApp report** to the group → menu → 📊 Daily WhatsApp Report

That is all. The rest of the sheets you only touch when something specific happens.

---

## 4. Recording a sale

**Go to the SALES LOG tab.**

Click the first empty row below the last entry and fill in each column:

| Column | What to enter |
|---|---|
| **Date** | Today's date — type it as DD/MM/YYYY, e.g. `28/02/2026` |
| **Customer Name** | The customer's name, or leave blank for a walk-in |
| **Product** | Click the cell — a dropdown appears — pick the product |
| **Qty** | How many units you sold |
| **Unit Price TZS** | This fills automatically from the price list. Change it only if you gave a discount |
| **Total TZS** | Fills automatically (Qty × Unit Price) — do not type here |
| **Amount Paid TZS** | How much the customer paid you right now |
| **Balance Owed TZS** | Fills automatically (Total − Paid) — do not type here |
| **Recorded By** | Click the cell — pick your name from the dropdown |
| **Notes** | Optional — e.g. "gave 500 TZS discount" or "paid half, owes rest" |

**If the balance owed is greater than 0**, the row will turn red automatically. This is a reminder that the customer still owes money. Also add them to the DEBTS tab (see section 5).

**If the customer paid in full**, the balance will be 0 and the row stays white — nothing else to do.

---

## 5. Recording a debt (customer owes money)

If a customer did not pay in full, record the debt so you can track it.

**Go to the DEBTS tab.**

Add a new row:

| Column | What to enter |
|---|---|
| **Date** | Date the debt started |
| **Customer Name** | Their name |
| **Items Description** | What they bought, e.g. "2× Mazoe Orange 2L, 1× Water 1.5L" |
| **Original Amount TZS** | The full amount they owed |
| **Amount Paid TZS** | How much they paid upfront (can be 0) |
| **Still Owed TZS** | Fills automatically |
| **Status** | Fills automatically — shows ⚠️ OUTSTANDING or ✅ CLEARED |
| **Recorded By** | Pick your name |
| **Notes** | Optional — e.g. "will pay Friday" |

---

## 6. Marking a debt as paid

When a customer comes back and pays:

1. Go to the **DEBTS** tab
2. Find their row
3. Update the **Amount Paid TZS** column — type the new total amount they have paid so far (not just the new payment, the full total)
4. The **Still Owed** and **Status** columns update automatically
5. When fully paid, the row turns green and shows ✅ CLEARED

**Also record the payment in SALES LOG** if you want it counted in today's revenue — add a new row with the product as something like "Debt Payment" or the original product, amount paid, and your name.

---

## 7. Recording an expense

Any money the business spends — transport, airtime, packaging, food for the stall, etc.

**Go to the EXPENSES tab.**

Add a new row:

| Column | What to enter |
|---|---|
| **Date** | Date of the expense |
| **Category** | Click — pick from dropdown (Transport, Packaging, Airtime, etc.) |
| **Description** | Brief description, e.g. "Taxi to market" or "Bubble wrap rolls" |
| **Amount** | The number only — no currency symbol |
| **Currency** | Click — pick **TZS** for Tanzania costs, **USD** for Zimbabwe costs |
| **Amount in TZS** | Fills automatically — USD amounts are converted using the exchange rate |
| **Paid By** | Pick your name |
| **Notes** | Optional |

**Important:** Stock purchase costs from Zimbabwe should be recorded in the STOCK RECEIVED tab — not here. Use EXPENSES for running costs only.

---

## 8. Logging new stock arriving from Zimbabwe

When stock arrives:

**Step 1 — Go to the STOCK RECEIVED tab** and add a new row:

| Column | What to enter |
|---|---|
| **Date** | Date the stock arrived |
| **Product** | Pick from dropdown |
| **Qty Received** | How many units arrived |
| **Cost Per Unit USD** | What you paid per unit in USD |
| **Total Cost USD** | Fills automatically |
| **Supplier / Notes** | Who you bought from, or any notes |
| **Recorded By** | Pick your name |

Do this for every product that arrived. If 3 different products came, add 3 rows.

**Step 2 — Update the STOCK tab** (see section 9 below).

---

## 9. Updating stock quantities

**Go to the STOCK tab.**

You only need to update two columns:

- **Current Qty** — update this whenever stock changes (after new stock arrives, or if you do a physical count)
- **Last Updated** — type today's date

The other columns (prices, values, low stock alerts) all update automatically.

**When does stock go down?** The STOCK tab does not automatically subtract stock when you record a sale — you update it manually when you do a count. Most small businesses update it when new stock arrives and when they do a weekly count.

**Low stock alerts:** If the Current Qty falls below the threshold set in SETTINGS, the row turns orange and the product appears in the Dashboard stock alerts section. If qty reaches 0, the row turns red.

---

## 10. Recording a partner withdrawal or contribution

If a partner takes money out of the business before the official profit split, or puts in extra money:

**Go to the PARTNER ACCOUNT tab** and add a row:

| Column | What to enter |
|---|---|
| **Date** | Date of the transaction |
| **Partner** | Pick the partner's name |
| **Type** | **Withdrawal** = they took money out / **Contribution** = they put money in |
| **Amount** | The number |
| **Currency** | TZS or USD |
| **USD Equivalent** | Fills automatically |
| **Notes** | What it was for, e.g. "took out for school fees" |

These adjustments are automatically factored into each partner's share when you refresh the Dashboard. A withdrawal reduces their share; a contribution increases it.

---

## 11. Refreshing the Dashboard

The Dashboard does not update itself — you have to refresh it manually.

1. Click the **🛒 ZimGoodies** menu at the top of the spreadsheet
2. Click **🔄 Refresh Dashboard**
3. Wait a few seconds
4. The DASHBOARD tab updates with all current figures

Do this at the end of each day, or any time you want to see the latest numbers.

---

## 12. Getting the daily WhatsApp report

1. Click the **🛒 ZimGoodies** menu
2. Click **📊 Daily WhatsApp Report**
3. A box pops up with today's report already formatted
4. Click the **📋 Copy to Clipboard** button
5. Open WhatsApp, go to your partners group, and paste

The report includes today's sales, cash collected, expenses, partner shares, stock alerts, and outstanding debts.

**Note:** The report is based on entries in SALES LOG that have today's date. Make sure you enter the correct date when recording sales.

---

## 13. Updating the exchange rate

Do this whenever the rate changes significantly.

1. Go to the **SETTINGS** tab
2. Find the row that says **💱 USD → TZS Exchange Rate**
3. Click the yellow cell (it shows the current rate, e.g. 3,200)
4. Type the new rate and press Enter

Every calculation in the whole spreadsheet updates immediately — profits, expenses, partner shares, everything.

---

## 14. Changing product prices

1. Go to the **SETTINGS** tab
2. Scroll down to the **📦 Product List** table
3. Find the product you want to update
4. Click the cost or price cell and type the new value
5. Press Enter

The STOCK sheet pulls prices from SETTINGS automatically — no need to update it separately.

---

## 15. Adding a new product

1. Go to the **SETTINGS** tab → **📦 Product List** table
2. Click the first empty row after the last product
3. Fill in: Product Name, Cost Price USD, Selling Price TZS, Low Stock Threshold
4. Go to the **STOCK** tab and add a new row for the product:
   - Type the product name in column A (must match SETTINGS exactly)
   - Type 0 in Current Qty
   - The price columns will auto-fill via VLOOKUP
5. To add the product to all dropdowns: click **🛒 ZimGoodies** menu → **⚙️ Setup / Reset Sheets** → Yes

---

## 16. Common questions

**Q: I made a mistake in a sale — can I fix it?**
Yes. Just click the cell with the wrong value and type the correct one. There are no locks. Be careful to also fix the DEBTS tab if the sale had a balance owed.

**Q: A customer's name is wrong — can I edit it?**
Yes, click the cell and retype it.

**Q: Can I delete a row?**
Yes. Right-click the row number on the left → Delete row. Only do this if you entered something by mistake — do not delete records of real transactions.

**Q: The Dashboard shows $0 for everything — why?**
Click **🛒 ZimGoodies** menu → **🔄 Refresh Dashboard**. The Dashboard only updates when you refresh it.

**Q: A formula cell is showing an error (like #REF! or #N/A) — what do I do?**
Do not edit formula cells. If you accidentally deleted a formula, you can get it back by running Setup again: **🛒 ZimGoodies** menu → **⚙️ Setup / Reset Sheets** → Yes. Your data is safe — setup only rebuilds formatting and formulas, it does not delete your entries.

**Q: Can two of us record sales at the same time?**
Yes. Google Sheets handles multiple editors at once. You might occasionally see a brief conflict message — just refresh the page and your entry will be saved.

**Q: What does "Recorded By" do?**
It just logs which partner entered the data. It helps if there is ever a question about a specific entry.

**Q: Can we add a new expense category?**
Yes. Go to **SETTINGS** and add the category to the expense categories list. It will appear in the dropdown automatically on the next row you enter.

**Q: The WhatsApp report shows 0 sales today but we sold things — why?**
Check that the sales in SALES LOG have today's date in the Date column in the format DD/MM/YYYY. If the date was entered differently (e.g. as text or in the wrong format) the report may not match it.

---

*ZimGoodies business tracker — built for Isabel, Fra & Rumbi.*
*GitHub: [https://github.com/isabelmoyo/zim_goodies](https://github.com/isabelmoyo/zim_goodies)*
