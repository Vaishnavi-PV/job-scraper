import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os

print("🚀 Script started...")

# URLs
base_url = "https://realpython.github.io"
url = base_url + "/fake-jobs/"

headers = {"User-Agent": "Mozilla/5.0"}

# Request page
response = requests.get(url, headers=headers)
soup = BeautifulSoup(response.text, "html.parser")

jobs = soup.find_all("div", class_="card-content")

job_list = []

# -------- Helper Functions -------- #

def get_experience(title):
    title = title.lower()
    if "senior" in title:
        return "5+ years"
    elif "lead" in title or "manager" in title:
        return "7+ years"
    elif "junior" in title:
        return "1-2 years"
    else:
        return "2-4 years"

def get_skills(title):
    title = title.lower()

    if "python" in title:
        return "Python, Django, Flask, SQL"
    elif "engineer" in title:
        return "Python, Java, SQL, Git"
    elif "developer" in title:
        return "JavaScript, React, HTML, CSS"
    elif "data" in title:
        return "Python, Pandas, Machine Learning, SQL"
    else:
        return "Communication, Problem Solving, Teamwork"

def get_description(title):
    return f"We are looking for a {title} who can contribute by building scalable solutions, collaborating with teams, and delivering high-quality applications."

# -------- Scraping -------- #

for job in jobs:
    title_tag = job.find("h2", class_="title")
    location_tag = job.find("p", class_="location")
    link_tag = job.find("a")

    title = title_tag.text.strip() if title_tag else "N/A"
    location = location_tag.text.strip() if location_tag else "N/A"
    link = base_url + link_tag["href"] if link_tag else "N/A"

    experience = get_experience(title)
    skills = get_skills(title)
    salary = "Not Disclosed"
    description = get_description(title)

    job_list.append({
        "JobTitle": title,
        "Location": location,
        "ExperienceRequired": experience,
        "SkillsRequired": skills,
        "Salary": salary,
        "JobURL": link,
        "JobDescriptionSummary": description
    })

# -------- Create DataFrame -------- #

df = pd.DataFrame(job_list)

# -------- Save Excel -------- #

file_name = "Final_Jobs.xlsx"

# Delete old file if exists (prevents permission error sometimes)
if os.path.exists(file_name):
    try:
        os.remove(file_name)
    except:
        print("⚠️ Close the Excel file before running again!")
        exit()

df.to_excel(file_name, index=False)

# -------- Format Excel -------- #

wb = load_workbook(file_name)
ws = wb.active

for column in ws.columns:
    max_length = 0
    col_letter = column[0].column_letter

    for cell in column:
        try:
            if cell.value:
                # Wrap text
                cell.alignment = Alignment(wrap_text=True)

                # Adjust width
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass

    ws.column_dimensions[col_letter].width = min(max_length + 5, 50)

wb.save(file_name)

print("✅ Done! Clean, structured, formatted Excel generated successfully.")