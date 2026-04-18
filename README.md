# Job Scraper Project

## * Description
This project is a Python-based web scraper that extracts job listings from a careers webpage and stores the data in a structured Excel file.

## * Features
- Scrapes job title, location, and job links
- Generates experience, skills, and job descriptions
- Saves data into a well-formatted Excel file
- Automatically adjusts column width and formatting

## * Technologies Used
- Python
- requests
- BeautifulSoup
- pandas
- openpyxl

## * How It Works
1. Sends an HTTP request to the website  
2. Parses HTML content using BeautifulSoup  
3. Extracts job details  
4. Generates additional fields (experience, skills, description)  
5. Stores data in a structured format  
6. Exports data to an Excel file  

## *  How to Run
1. Install required libraries:pip install requests beautifulsoup4 pandas openpyxl
2. Run the script:python scraper.py


## * Output
The program generates:

- `Final_Jobs.xlsx` → Contains structured and formatted job data

## * Source Website
https://realpython.github.io/fake-jobs/

## * Author
Vaishnavi PV
