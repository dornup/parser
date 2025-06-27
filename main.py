from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup
from lxml.html.soupparser import fromstring

last_art = 100088
last_row = 308
table = load_workbook("C:\\Users\\Admin\\Desktop\\vsc_projects\\parser\\parser_excel.xlsx")
sheet = table["Лист1"]
cathegory_page = requests.get("https://xn--80agf5ao.xn--p1acf/%d0%b3%d0%bb%d0%b0%d0%b2%d0%bd%d0%b0%d1%8f/%d0%bf%d1%80%d0%be%d0%b4%d1%83%d0%ba%d1%86%d0%b8%d1%8f/%d0%b8%d0%b7%d0%be%d1%82%d0%b5%d1%80%d0%bc%d0%b8%d1%87%d0%b5%d1%81%d0%ba%d0%b8%d0%b5-%d0%b8-%d1%81%d0%bb%d1%83%d0%b6%d0%b5%d0%b1%d0%bd%d1%8b%d0%b5-%d0%b4%d0%b2%d0%b5%d1%80%d0%b8/") # TODO ссылка категории, ее будем заменять
elem = fromstring(str(BeautifulSoup(cathegory_page.text, 'html.parser')))
n = 5 # TODO
pages = [elem.xpath(f'//*[@id="content"]/div/div/div[{3 + 2*(i//3) if i%3 != 0 else 3 + 2*(i//3 - 1)}]/div/div[{i%3 if i%3!=0 else 3}]/div[1]/{"div/figure" if i == 13 else "div"}/a/@href')[0] for i in range(1, n + 1)] # скорее всего тоже будет видоизменяться
cathegory = elem.xpath('//*[@id="content"]/div/div/div[1]/div/div[1]/div[1]/div/h2/text()')[0].strip(" -")
# undercathegory = elem.xpath('//*[@id="content"]/div/div/div[1]/div/div[1]/div[2]/div/h2/text()')[0].strip()

for j, page in enumerate(pages):
    last_row +=1
    site = requests.get(page)
    s_elem = fromstring(str(BeautifulSoup(site.text, 'html.parser')))
    typ = "variable"
    sheet[f'B{last_row}'].value = typ
    last_art += 1
    sheet[f'C{last_row}'].value = str(last_art)
    name = s_elem.xpath('//*[@id="content"]/div/div/div[1]/div/div[2]/div[1]/div/h2/text()')[0]
    sheet[f'D{last_row}'].value = name
    short_description = name.capitalize().join(s_elem.xpath('//*[@id="content"]/div/div/div[2]/div/div[2]/div[1]/div/p/text()'))
    # //*[@id="content"]/div/div/div[2]/div/div[2]/div[2]/div/p
    sheet[f'E{last_row}'].value = short_description
    sheet[f'L{last_row}'].value = cathegory
    # sheet[f'M{last_row}'].value = undercathegory
    img = s_elem.xpath(f'//*[@id="content"]/div/div/div[1]/div/div[1]/div/{"div/figure" if j == 13 else "div"}/img/@src')[0]
    # //*[@id="content"]/div/div/div[1]/div/div[1]/div/div/img
    sheet[f'N{last_row}'].value = img
    atr_name = "Модель"
    sheet[f'Q{last_row}'].value = atr_name
    atr_value = []
    for i in s_elem.xpath('///*[@class="e-n-tab-title elementor-animation-shrink"]/span/text()'):
        s = str(i.strip())
        atr_value.append(s)
    sheet[f'R{last_row}'].value = ", ".join(atr_value)
    #kids, hahaha

    for i, elem in enumerate(atr_value):
        print(j)
        last_row += 1
        typ_k = 'variation'
        sheet[f'B{last_row}'].value = typ_k
        art = f'{last_art}-{i+1}'
        sheet[f'C{last_row}'].value = art
        name_k = elem
        sheet[f'D{last_row}'].value = name_k
        temperature = s_elem.xpath('//*[@role="tabpanel"]/div/div[2]/div/div[2]/div[1]/div[2]/div/h2/text()')[i]
        sheet[f'G{last_row}'].value = temperature
        tiporazmer = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[2]/div/div[2]/div[1]/div[4]/div/h2/text()') + s_elem.xpath('//*[@id="e-n-tab-content-13646098364"]/div[2]/div/div[2]/div[1]/div[4]/div/h2/text()'))[i]
        sheet[f'H{last_row}'].value = tiporazmer
        vikladka = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[2]/div/div[2]/div[1]/div[6]/div/h2/text()') + s_elem.xpath('//*[@role="tabpanel"]/div[2]/div/div[2]/div[1]/div[6]/div/h2/text()'))[i]
        sheet[f'I{last_row}'].value = vikladka
        obem = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[2]/div/div[2]/div[1]/div[8]/div/h2/text()')+ s_elem.xpath('//*[@role="tabpanel"]/div[2]/div/div[2]/div[1]/div[8]/div/h2/text()'))[i]
        sheet[f'J{last_row}'].value = obem
        holodosnabj = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[2]/div/div[2]/div[1]/div[10]/div/h2/text()')+ s_elem.xpath('//*[@role="tabpanel"]/div[2]/div/div[2]/div[1]/div[10]/div/h2/text()'))[i]
        sheet[f'K{last_row}'].value = holodosnabj
        sheet[f'L{last_row}'].value = cathegory
        # sheet[f'M{last_row}'].value = undercathegory
        img1 = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[1]/div/div[1]/div/img/@src') + s_elem.xpath('//*[@role="tabpanel"]/div[1]/div/div[1]/div/img/@src'))[i]
        sheet[f'N{last_row}'].value = img1
        img2 = list(s_elem.xpath('//*[@role="tabpanel"]/div/div[1]/div/div[2]/div/img/@src') + s_elem.xpath('//*[@role="tabpanel"]/div[1]/div/div[2]/div/img/@src'))[i]
        sheet[f'O{last_row}'].value = img2
        rod = last_art
        sheet[f'P{last_row}'].value = rod
        atr_name_k = atr_name
        sheet[f'Q{last_row}'].value = atr_name_k
        atr_value_k = elem
        sheet[f'R{last_row}'].value = atr_value_k

table.save("C:\\Users\\Admin\\Desktop\\vsc_projects\\parser\\parser_excel.xlsx")
table.close()