# 📊 EPAM Monthly Certification Report Tool

This Python-based tool automates the generation of certification analytics from Excel data. It processes employee certification records and produces:

- 📄 A clean Excel output with structured sheets.
- 📈 Multiple insightful visual charts.
- 📘 A complete PDF report, including trends and statistics.
- 🌐 Real-time web scraping of trending GCP & Azure certifications.
- ✅ Status validation using Issue and Expiry Dates.
- 📊 Distribution of certification levels (Fundamentals, Associate, Expert).
- 👤 Coverage of certified vs non-certified members in the team.

---

## 📁 Project Structure

```
.
├── monthly_cert_report.py       # Main script
├── data.xlsx                    # Input Excel file (example)
├── certification_report.xlsx    # Output Excel file
├── Certification_Report.pdf     # Generated PDF report
├── charts/                      # Folder containing all generated chart PNGs
└── epam_logo.png                # Optional logo for the PDF report
```

---

## 🚀 How to Run

### 1. ✅ Install Requirements

```bash
pip install instructions.txt
```

### 2. ▶️ Execute the Script

```bash
python monthly_cert_report.py data.xlsx certification_report.xlsx
```

This will:
- Read the `data.xlsx` input
- Generate visual charts
- Create a structured Excel file and a PDF report

---

## 📥 Input Format (data.xlsx)

The `Export` sheet must include at least the following columns:

- `Employee Name`
- `Certificate Name`
- `Program Title`
- `Issue Date`
- `Expiry Date`
- `Track` (optional)
- `Primary Skill` (optional)

---

## 📌 Charts Generated

- **Certification Level Distribution**
- **Certification Status Distribution**
- **Trending Certification Market Share**
- **Certified vs Not Certified Users in the Unit**

Each chart is saved in the `charts/` folder and included in the final PDF report.

---

## 🌐 Web-Scraped Trending Certifications

Trending GCP and Azure certifications are pulled from:

- [Google Cloud Certification Page](https://cloud.google.com/certification)
- [Microsoft Learn Certifications](https://learn.microsoft.com/en-us/certifications/)

These are categorized as **Cloud** or **AI**, and visualized in the PDF.

---

## 🧠 Logic Highlights

- Certificates without an expiry date are considered **never expiring**.
- Status is determined by comparing expiry dates to today's date.
- Certifications are categorized by level: Fundamentals, Associate, Expert.
- A bar chart summarizes how many of the **51 team members** are certified.

---

## 🏷️ License

MIT License.

---

## 🤝 Maintainer

Developed and maintained by the EPAM GCP & Azure Leadership team.
