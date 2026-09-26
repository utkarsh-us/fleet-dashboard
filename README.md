# 🚛 Steelworks Fleet Compliance Dashboard

A modern, interactive Streamlit dashboard for tracking fleet document compliance — **RC, Insurance, Fitness, MV Tax, Permit, and TP** — across **SWPE** and **SWIN** entities.

**Live App:** [steelworks-fleet.streamlit.app](https://steelworks-fleet.streamlit.app)

---

## ✨ Features

- **Entity Toggle** — Switch between SWPE and SWIN fleet data
- **Full-Text Search** — Search across **all columns** (Vehicle No, Name, Chassis, Engine, Policy No, RTO, etc.)
- **Document Filters** — View compliance for Insurance, Fitness, MV Tax, Permit, TP
- **Year & Month Filters** — Narrow down to a specific period
- **KPI Cards** — Total, Active, Expiring Soon, Expired — live counts
- **📅 Monthly Due Date Breakdown** — Click any month to see vehicles whose documents are due
- **Compact Records Table** — Clean per-row layout with **🔍 View** button in the last column
- **Vehicle Drill-Down Popup** — Click View to see full vehicle details + one-click access to each document
- **Color-Coded Status** — 🟢 Active · 🟡 Expiring (≤30 days) · 🔴 Expired · ⚪ N/A
- **Direct Document Access** — Opens SharePoint files/folders directly from the dashboard
- **Last Updated Timestamp** — Shows when the source Excel was last modified
- **Modern UI** — Gradient headers, card-based KPIs, clean typography

---

## 🖥️ How to Use

1. Open the dashboard: **[steelworks-fleet.streamlit.app](https://steelworks-fleet.streamlit.app)**
2. Use the **sidebar filters** to:
   - Search any value (vehicle no, chassis, policy no, RTO, etc.)
   - Select the **Entity** (SWPE / SWIN)
   - Choose the **Document Type** (Insurance / Fitness / MV Tax / Permit / TP)
   - Optionally narrow by **Year** and **Month**
3. Review the **KPI cards** at the top for a quick summary
4. Scroll to **📅 Monthly Due Date Breakdown** — click any month card to see the vehicle list with due dates
5. Scroll to **📋 Records** — click **🔍 View** on any row to open the vehicle popup
6. In the popup, click **📄 View [Document]** to open the PDF in a new tab

---

## 📁 Project Structure
