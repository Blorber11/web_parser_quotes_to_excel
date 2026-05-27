# Парсер цитат с сайта ://toscrape.com
# Собирает текст цитаты, автора и теги, сохраняет данные в Excel.
# Использованные библиотеки: BeautifulSoup4, Requests, Pandas, Openpyxl.
from bs4 import BeautifulSoup
import requests
import pandas as pd

url = 'https://quotes.toscrape.com/'
response = requests.get(url)
# ответ от сайта
#print(response)

bs = BeautifulSoup(response.text,"lxml") # html код 

quotes = bs.find_all("div", class_="quote")
data = []

for block in quotes:
    # текст цитаты внутри блока
    text = block.find("span", class_="text").text

    # автор внутри этого блока
    author = block.find("small", class_="author").text

    # блок с тегами
    tags_block = block.find("div", class_="tags")

    # все ссылки <a> внутри этого блока тегов
    tags_links = tags_block.find_all("a", class_="tag")

    # все теги этой цитаты в одну строку через запятую
    tags_list = []
    for tag_link in tags_links:
        tags_list.append(tag_link.text)
      
    tags_string = ", ".join(tags_list)
    
    # упаковываю данные одной цитаты в словарь и добавляю в общий список
    data.append(
        {"Цитата": text, "Автор": author, "Теги": tags_string}
    )

table = pd.DataFrame(data)
table.to_excel("quotes_portfolio.xlsx", index=False) # создастся таблица эксель с цитатами
