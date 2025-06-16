# monthly_cert_report.py

# 📊 Requirements:
# Make sure to install the following libraries before running:
# pip install pandas matplotlib seaborn fpdf openpyxl xlsxwriter beautifulsoup4 requests

# ✅ Optional: Apply a matplotlib style for better visuals
import matplotlib.pyplot as plt
plt.style.use("seaborn-v0_8-darkgrid")
import pandas as pd
import matplotlib.ticker as ticker
import seaborn as sns
import sys
import os
import requests
from bs4 import BeautifulSoup
from fpdf import FPDF
from datetime import datetime

# Define a color palette
COLOR_PALETTE = sns.color_palette("Set2") + sns.color_palette("Set3") + sns.color_palette("pastel")

# Reference list of most demanded GCP/Azure certifications with market percentage
MOST_DEMANDED_CERTS = [
    ("Microsoft", "AZ-104", "Cloud", 12.5),
    ("Microsoft", "AZ-305", "Cloud", 10.3),
    ("Microsoft", "AI-900", "AI", 8.7),
    ("Microsoft", "AI-102", "AI", 7.2),
    ("Google", "Professional Cloud Architect", "Cloud", 15.6),
    ("Google", "Associate Cloud Engineer", "Cloud", 14.8),
    ("Google", "Professional Data Engineer", "AI", 11.4),
    ("Google", "Professional Machine Learning Engineer", "AI", 9.1),
    ("Google", "Generative AI Leader", "AI", 6.2)
]

def generate_level_distribution_chart(df_export, output_folder):
    def detect_level(name):
        name = str(name).lower()

        # Expert/Professional
        if any(keyword in name for keyword in [
            'az-400', 'professional', 'expert'
        ]):
            return 'Expert/Professional'

        # Associate
        elif any(keyword in name for keyword in [
            'cka',
            'associate',
            'sitecore 10 system administrator',
            'az-305'
        ]):
            return 'Associate'

        # Fundamentals
        elif any(keyword in name for keyword in [
            'fundamental',
            'foundation',
            'aws certified cloud practitioner',
            'cloud digital leader',
            'linux essentials',
            'le-1',
            'scrum foundation professional certificate',  # SFPC
            'sfpc'
        ]):
            return 'Fundamentals'

        else:
            return 'Other'

    df_export['Level'] = df_export['Certificate Name'].apply(detect_level)
    level_counts = df_export['Level'].value_counts()
    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(level_counts.index, level_counts.values, color=sns.color_palette("muted"))
    ax.set_title("Certification Level Distribution", fontsize=12, fontweight='bold')
    ax.set_ylabel("Count", fontsize=10)
    ax.set_xlabel("Level", fontsize=10)
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, str(bar.get_height()), ha='center', fontsize=9)

    plt.tight_layout()
    os.makedirs(os.path.join(output_folder, "charts"), exist_ok=True)
    chart_path = os.path.join(output_folder, "charts", "Certification_Level_Distribution.png")
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"Chart saved: {chart_path}")



def scrape_trending_certifications():
    try:
        gcp_url = "https://cloud.google.com/certification"
        azure_url = "https://learn.microsoft.com/en-us/certifications/"

        gcp_res = requests.get(gcp_url)
        azure_res = requests.get(azure_url)

        gcp_soup = BeautifulSoup(gcp_res.text, "html.parser")
        azure_soup = BeautifulSoup(azure_res.text, "html.parser")

        certs = []

        for card in gcp_soup.select(".devsite-landing-row-item-title"):
            name = card.text.strip()
            if any(x in name.lower() for x in ["cloud", "ai", "data"]):
                category = "AI" if "ai" in name.lower() or "machine" in name.lower() else "Cloud"
                certs.append(("Google", name, category, None))

        for card in azure_soup.select("a.card-title"):
            name = card.text.strip()
            if any(x in name.lower() for x in ["azure", "ai", "cloud"]):
                category = "AI" if "ai" in name.lower() or "data" in name.lower() else "Cloud"
                certs.append(("Microsoft", name, category, None))

        return certs if certs else MOST_DEMANDED_CERTS
    except Exception as e:
        print(f"Error fetching trending certifications: {e}")
        return MOST_DEMANDED_CERTS

def generate_certified_ratio_chart(df_export, total_employees, output_folder):
    certified_count = df_export['Employee Name'].nunique()
    uncertified_count = total_employees - certified_count

    labels = ['Certified', 'Not Certified']
    values = [certified_count, uncertified_count]
    colors = ['#66c2a5', '#fc8d62']

    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(labels, values, color=colors)
    ax.set_title("Certification Coverage in the Unit", fontsize=12, fontweight='bold')
    ax.set_ylabel("Number of People", fontsize=10)

    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.5,
                str(bar.get_height()), ha='center', fontsize=9)

    plt.tight_layout()
    chart_path = os.path.join(output_folder, "charts", "Certification_Coverage_Bar.png")
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"Chart saved: {chart_path}")


def generate_trending_cert_table_sheet(writer):
    df_trending = pd.DataFrame(MOST_DEMANDED_CERTS, columns=["Provider", "Certification Name", "Category", "Market %"])
    df_trending.to_excel(writer, sheet_name='Trending Certs', index=False)

def generate_trending_chart(trending_data, output_folder):
    df = pd.DataFrame(trending_data, columns=["Provider", "Certification Name", "Category", "Market %"])
    df = df.dropna(subset=["Market %"])
    df.sort_values("Market %", ascending=True, inplace=True)
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(df["Certification Name"], df["Market %"], color='#6fa8dc', edgecolor='black')
    ax.set_title("Market Share of Top Certifications", fontsize=16, fontweight='bold')
    ax.set_xlabel("% Market", fontsize=12)
    ax.grid(axis='x', linestyle='--', linewidth=0.5, alpha=0.7)
    ax.tick_params(axis='y', labelsize=8)

    for bar in bars:
        ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height()/2, f"{bar.get_width():.1f}%", va='center', fontsize=9)

    plt.tight_layout()
    os.makedirs(os.path.join(output_folder, "charts"), exist_ok=True)
    chart_path = os.path.join(output_folder, "charts", "Trending_Market_Share.png")
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"Chart saved: {chart_path}")

def generate_status_chart(df_export, output_folder):
    status_counts = df_export['Status'].value_counts()
    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(status_counts.index, status_counts.values, color=sns.color_palette("pastel"))
    ax.set_title("Certification Status Distribution", fontsize=12, fontweight='bold')
    ax.set_ylabel("Count", fontsize=10)
    ax.set_xlabel("Status", fontsize=10)
    for bar in bars:
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1, str(bar.get_height()), ha='center', fontsize=9)

    plt.tight_layout()
    chart_path = os.path.join(output_folder, "charts", "Certification_Status_Bar.png")
    plt.savefig(chart_path, dpi=300)
    plt.close()
    print(f"Chart saved: {chart_path}")

def generate_pdf_report(output_folder, logo_path=None):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)

    chart_folder = os.path.join(output_folder, "charts")
    charts = [f for f in os.listdir(chart_folder) if f.endswith(".png")]

    charts.sort()
    if "Certification_Level_Distribution.png" in charts:
        charts.insert(0, charts.pop(charts.index("Certification_Level_Distribution.png")))

    for i in range(0, len(charts), 2):
        pdf.add_page()
        if logo_path and os.path.exists(logo_path):
            pdf.image(logo_path, x=10, y=8, w=25)
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(0, 10, "EPAM GCP & Azure Certification", ln=True, align='C')
        pdf.ln(10)
        if i < len(charts):
            pdf.image(os.path.join(chart_folder, charts[i]), x=10, y=30, w=180)
        if i + 1 < len(charts):
            pdf.image(os.path.join(chart_folder, charts[i + 1]), x=10, y=150, w=180)

    pdf.add_page()
    if logo_path and os.path.exists(logo_path):
        pdf.image(logo_path, x=10, y=8, w=25)
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Trending Certifications in GCP & Azure (2025)", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(50, 10, "Provider", 1, 0, 'C')
    pdf.cell(90, 10, "Certification Name", 1, 0, 'C')
    pdf.cell(30, 10, "Category", 1, 0, 'C')
    pdf.cell(20, 10, "% Market", 1, 1, 'C')
    pdf.set_font("Arial", '', 10)
    for vendor, cert, category, market in MOST_DEMANDED_CERTS:
        pdf.cell(50, 10, vendor, 1)
        pdf.cell(90, 10, cert[:40], 1)
        pdf.cell(30, 10, category, 1)
        pdf.cell(20, 10, f"{market:.1f}%" if market else "-", 1)
        pdf.ln()

    pdf.add_page()
    month_label = datetime.now().strftime("Report generated in %B %Y")
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, month_label, ln=True, align='C')

    pdf_output = os.path.join(output_folder, "Certification_Report.pdf")
    pdf.output(pdf_output)
    print(f"PDF report created: {pdf_output}")

def generate_certification_report(input_path, output_path):
    df_export = pd.read_excel(input_path, sheet_name='Export')
    output_dir = os.path.dirname(output_path) or '.'
    logo_path = os.path.join(output_dir, 'epam_logo.png')

    # Generar gráfico de niveles de certificación
    generate_level_distribution_chart(df_export, output_dir)

    # Generar el PDF con el gráfico ya disponible
    generate_pdf_report(output_dir, logo_path=logo_path)
    print(f"Report saved to: {output_path}")
    generate_certified_ratio_chart(df_export, total_employees=51, output_folder=output_dir)


    today = pd.to_datetime(datetime.today().date())
    df_export['Issue Date'] = pd.to_datetime(df_export['Issue Date'], errors='coerce')
    df_export['Expiry Date'] = pd.to_datetime(df_export['Expiry Date'], errors='coerce')
    df_export["Category"] = df_export.apply(
    lambda row: classify_certificate(row["Certificate Name"], row["Program Title"]), axis=1)
    df_export['Status'] = df_export['Expiry Date'].apply(
        lambda x: 'Active' if pd.isna(x) or x > today else 'Expired'
    )

    cert_df = df_export[['Employee Name', 'Certificate Name', 'Program Title', 'Category', 'Issue Date', 'Expiry Date', 'Status']].drop_duplicates()
    cert_df = pd.merge(
        cert_df,
        df_export[['Employee Name', 'Certificate Name', 'Track']].drop_duplicates(),
        on=['Employee Name', 'Certificate Name'], how='left'
    )
    cert_df = pd.merge(
        cert_df,
        df_export[['Employee Name', 'Certificate Name', 'Primary Skill']].drop_duplicates(),
        on=['Employee Name', 'Certificate Name'], how='left'
    )

    cert_df = cert_df[['Employee Name', 'Certificate Name', 'Program Title', 'Category', 'Track', 'Primary Skill', 'Issue Date', 'Expiry Date', 'Status']]

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, sheet_name='Export', index=False)
        cert_df.to_excel(writer, sheet_name='Certy by Category', index=False)
        generate_trending_cert_table_sheet(writer)

    output_dir = os.path.dirname(output_path) or '.'
    logo_path = os.path.join(output_dir, 'epam_logo.png')
    generate_trending_chart(MOST_DEMANDED_CERTS, output_dir)
    generate_status_chart(df_export, output_dir)
    generate_pdf_report(output_dir, logo_path=logo_path)
    print(f"Report saved to: {output_path}")

def classify_certificate(cert_name, program_title):
    name = str(cert_name).lower()
    vendor = str(program_title).lower()

    if vendor == 'microsoft':
        if 'az-500' in name:
            return 'Security Cloud'
        elif 'ai-900' in name or 'azure ai fundamentals' in name:
            return 'AI'
        elif 'dp-900' in name or 'azure data fundamentals' in name:
            return 'AI'
        elif 'pl-900' in name or 'power platform' in name:
            return 'Methodology'
        elif 'az-700' in name or 'network engineer' in name:
            return 'Cloud'
        elif 'az-303' in name or 'az-305' in name or 'azure architect' in name:
            return 'Cloud'
        elif 'administrator' in name or 'developer' in name:
            return 'Cloud'
        elif 'devops' in name:
            return 'DevOps'
        elif 'fundamentals' in name:
            return 'Cloud'

    if 'associate data practitioner' in name:
        return 'AI'

    if 'cloud' in name or 'solutions architect' in name:
        return 'Cloud'
    elif 'ai' in name or 'machine learning' in name or 'data engineer' in name or 'data fundamentals' in name:
        return 'AI'
    elif 'devops' in name or 'terraform' in name or 'cka' in name or 'sysops' in name:
        return 'DevOps'
    elif 'security' in name:
        return 'Security Cloud'
    elif 'itil' in name or 'scrum' in name or 'sfpc' in name:
        return 'Methodology'
    elif 'vmware' in name or 'linux' in name or 'juniper' in vendor or 'sitecore' in vendor or 'exam 740' in name:
        return 'Infrastructure'
    else:
        return 'Other'

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python monthly_cert_report.py <input_excel_file> <output_excel_file>")
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        if not os.path.exists(input_file):
            print(f"Error: File '{input_file}' does not exist.")
        else:
            generate_certification_report(input_file, output_file)
            print("Report generation completed.")
            print(f"Report saved to: {output_file}")