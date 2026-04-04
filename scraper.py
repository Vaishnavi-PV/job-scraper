import requests
from bs4 import BeautifulSoup
import pandas as pd

# Step 1: Website link
url = "https://realpython.github.io/fake-jobs/"

# Step 2: Get webpage
response = requests.get(url)

# Step 3: Convert to readable format
soup = BeautifulSoup(response.text, "html.parser")

# Step 4: Find job cards
jobs = soup.find_all("div", class_="card-content")

job_list = []

# Step 5: Extract data
for job in jobs:
    title = job.find("h2", class_="title").text.strip()
    company = job.find("h3", class_="company").text.strip()
    location = job.find("p", class_="location").text.strip()

    job_list.append({
        "Title": title,
        "Company": company,
        "Location": location
    })

# Step 6: Save to Excel
df = pd.DataFrame(job_list)
df.to_excel("jobs.xlsx", index=False)

print("Done! File saved as jobs.xlsx")