import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://realpython.github.io/fake-jobs/"
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")

jobs = soup.find_all("div", class_="card-content")

job_list = []

for job in jobs:
    title = job.find("h2", class_="title").text.strip()
    location = job.find("p", class_="location").text.strip()
    job_url = "https://realpython.github.io" + job.find("a")["href"]

    job_list.append({
        "JobTitle": title,
        "Location": location,
        "ExperienceRequired": "",
        "SkillsRequired": "",
        "Salary": "",
        "JobURL": job_url,
        "JobDescriptionSummary": ""
    })

columns = [
    "JobTitle",
    "Location",
    "ExperienceRequired",
    "SkillsRequired",
    "Salary",
    "JobURL",
    "JobDescriptionSummary"
]

df = pd.DataFrame(job_list, columns=columns)
df.to_excel("Company_Jobs.xlsx", index=False)

print("Done! File saved as Company_Jobs.xlsx")
