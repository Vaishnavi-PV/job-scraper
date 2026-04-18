import requests
from bs4 import BeautifulSoup
import pandas as pd
# Base URL
base_url = "https://realpython.github.io"
url = base_url + "/fake-jobs/"
# Get main page
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")
# Find job cards
jobs = soup.find_all("div", class_="card-content")
job_list = []

# Loop through each job
for job in jobs:
    # Safe extraction
    title_tag = job.find("h2", class_="title")
    location_tag = job.find("p", class_="location")
    link_tag = job.find("a")

    title = title_tag.text.strip() if title_tag else ""
    location = location_tag.text.strip() if location_tag else ""
    job_url = ""
    if link_tag and "href" in link_tag.attrs:
        href = link_tag["href"].strip()
        if href.startswith("http"):
            job_url = href
        else:
            job_url = base_url + href
    # Visit job detail page
    summary = ""
    skills = ""
    if job_url:
        try:
            job_page = requests.get(job_url)
            job_soup = BeautifulSoup(job_page.text, "html.parser")
            # Description
            desc_tag = job_soup.find("div", class_="content")
            if desc_tag:
                description = desc_tag.get_text(separator=" ", strip=True)
                summary = description[:200]
            # Skills
            skills_tags = job_soup.find_all("li")
            if skills_tags:
                skills = ", ".join([s.get_text(strip=True) for s in skills_tags[:5]])

        except:
            summary = ""
            skills = ""

    # Append data
    job_list.append({
        "JobTitle": title,
        "Location": location,
        "ExperienceRequired": "",
        "SkillsRequired": skills,
        "Salary": "",
        "JobURL": job_url,
        "JobDescriptionSummary": summary
    })

# Required columns
columns = [
    "JobTitle",
    "Location",
    "ExperienceRequired",
    "SkillsRequired",
    "Salary",
    "JobURL",
    "JobDescriptionSummary"
]

# Create DataFrame
df = pd.DataFrame(job_list, columns=columns)

# Clean missing values
df = df.fillna("")

# Save Excel
df.to_excel("Company_Jobs.xlsx", index=False)

print("✅ Done! Clean Excel generated successfully.")