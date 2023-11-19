# biblioteki
import requests
import lxml
from bs4 import BeautifulSoup
import openpyxl
#bazovie danie
user = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0"
headers = {"User-agent" : user}
sess = requests.Session()
counter = 0
book = openpyxl.Workbook()
sheet = book.active
file = open('output.txt', "a", encoding = "utf-8")
#chistka
f = open('catalog.xlsx', 'r+')
f.truncate(0)
f.close()
print('Starting code')
#kod
for j in range (1,50):
  url = f"https://kups.club/?page={j}/"
  resp = sess.get(url, headers = headers)
  soup = BeautifulSoup(resp.text, "html.parser")
  products = soup.findAll("div", class_ = 'col-lg-4 col-md-4 col-sm-6 portfolio-item')
  for product in products:
    counter += 1
    title = product.find("h3", class_ = "card-title").text
    price = product.find("p", class_ = "card-text").text
    try:
      sponsor_tag = product.find("a", class_="text-black link-default")
      if sponsor_tag:
          sponsor_text = sponsor_tag.text.strip()
          sponsor_img = sponsor_tag.find('img', alt=True)
          sponsor = sponsor_img['alt'].strip() if sponsor_img and 'alt' in sponsor_img.attrs else sponsor_text
      else:
          sponsor = None
    except Exception as e:
      print("error")
      sponsor = None
    sheet[f"A{counter}"] = title
    sheet[f"B{counter}"] = price
    sheet[f"C{counter}"] = sponsor.strip() if sponsor else None
    file.write(f"Product No {counter}\n")
    file.write(f"Name: {title} {price} Sponsor: {sponsor}\n")
book.save('catalog.xlsx')
book.close()
print('Done')