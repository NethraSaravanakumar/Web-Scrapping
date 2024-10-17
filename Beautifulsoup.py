from bs4 import BeautifulSoup
import requests

import openpyxl

excel = openpyxl.Workbook()

Sheet = excel.active
Sheet.title="Title1"
print(excel.sheetnames)

Sheet.append(['Name1','Name2','Name3','Name4'])

try:
  source = requests.get('Enter a website')
  source.raise_for_status()

  soup = BeautifulSoup(source.text,'html.parser')
  #print(soup)

  songs = soup.find('div',class_="lister-list").find_all('div',class_="lister-item mode-advanced")
  #print(len(movies))
  for song in songs:
    name = song.find('h3',class_="lister-item-header").a.text
    rank = song.find('h3',class_="lister-item-header").get_text(strip=True).split('.')[0]
    year = song.find('h3',class_="lister-item-header").get_text(strip=False).split('\n')[3].strip('()')
    rating = song.find('div',class_="ratings-bar").strong.text
    Sheet.append([rank,name,year,rating])

except Exception as e:
  print(e)

excel.save('Top Movies.xlsx')
