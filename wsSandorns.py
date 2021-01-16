from requests_html import HTMLSession
from bs4 import BeautifulSoup
import xlwings as xl
from tqdm import tqdm

downloadedData = []

for i in tqdm(range(1, 7)):
  url = f'https://www.sanborns.com.mx/categoria/130102/el/{i}/f_filtros/W10=/'
  session = HTMLSession()
  contenedor = session.get(url).html
  htmlCode = BeautifulSoup(contenedor.html, 'html.parser')
  address = 'article.productbox div.carruselContenido'
  productContainer = htmlCode.select(address)

  for x in productContainer:
    link = f'https://www.sanborns.com.mx{x.find("a").attrs["href"]}'
    productName = x.find("a").select_one('p').text
    price = x.select_one('div.info span.preciodesc').text
    downloadedData.append(tuple([
      productName,
      price,
      link
    ]))

headers = ["Product", "Price","Link"]

wb = xl.Book()
ws = wb.sheets[0]

xl.books.active
xl.sheets.active

ws.range((1,1)).value = headers
ws.range((2,1)).value = downloadedData


