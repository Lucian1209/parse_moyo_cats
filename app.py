import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# Set the user-agent to a popular web browser to avoid being blocked
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}

# Send an HTTP request to the website and get its HTML content
url = 'https://www.moyo.ua/sitemap.html'
response = requests.get(url, headers=headers)
html_content = response.content

# Parse the HTML content using BeautifulSoup
soup = BeautifulSoup(html_content, 'html.parser')

# Find the product category elements in the parsed HTML using BeautifulSoup's find_all() method
product_categories = soup.find_all('a')

# Create a new Excel workbook and get the active sheet
workbook = Workbook()
sheet = workbook.active

# Write a header row to the sheet
sheet.append(['Product Category', 'Link'])

# Extract the relevant information from the product category elements and write them to the sheet
for category in product_categories:
    category_name = category.text.strip()
    category_link = category['href']
    sheet.append([category_name, category_link])

# Save the workbook to a file
workbook.save('moyo_categories.xlsx')

