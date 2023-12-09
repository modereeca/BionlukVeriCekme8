import openpyxl
from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup

wb = Workbook()
ws = wb.active
url = "https://www.rehberfx.com/arama?query=izmir&page="
basliklar = ("FİRMA ADI", "CEP NUMARASI", "TEL", "MAİL ADRESİ", "URL")
a = 1

for baslik in basliklar:
    ws.cell(column=a, row=1).value = baslik
    a += 1

query = "#gsc.tab=0&gsc.q=izmir&gsc.page=1"
r = 2
Offset = 3959
while Offset < 6328:
    webpage_main = requests.get(url + str(Offset) + query)
    soup_main = BeautifulSoup(webpage_main.content, "html.parser")
    ilanlar = soup_main.select(".pull-left.thumbnail")
    for i in range(len(ilanlar)):
        ilan_linki = ilanlar[i].attrs.get("href")
        print(ilan_linki)

        webpage = requests.get(ilan_linki)
        soup = BeautifulSoup(webpage.content, "html.parser")

        tel = soup.select(".fa.fa-phone.fa-fw")
        mobil_tel = soup.select(".fa.fa-mobile-phone.fa-fw")
        email = soup.select(".fa.fa-envelope.fa-fw")
        firmaadi = soup.select(".media-heading.dbox-title")

        firmaadi2 = firmaadi[0].find('small').replace_with('')



        if firmaadi is not None:
            ws.cell(column=1, row=r).value = firmaadi[0].text.lstrip().rstrip()

        if len(mobil_tel) > 0:
            ws.cell(column=2, row=r).value = mobil_tel[0].parent.find_next_sibling().text

        if len(tel) > 0:
            ws.cell(column=3, row=r).value = tel[0].parent.find_next_sibling().text

        if len(email) > 0:
            ws.cell(column=4, row=r).value = email[0].parent.find_next_sibling().text

        ws.cell(column=5, row=r).value = ilan_linki
        r += 1

    Offset += 1
    print(Offset)
    wb.save("firma_bilgileri2.xlsx")