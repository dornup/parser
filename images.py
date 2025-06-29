from openpyxl import load_workbook
import requests

wb = load_workbook("C:\\Users\\Admin\\Desktop\\vsc_projects\\parser\\parser_excel.xlsx")
ws = wb.worksheets[0]

links_1 = [row[13].value for row in ws]
links_1.pop(0)
links_2 = [row[14].value for row in ws]
links_2.pop(0)

for number, link_group in enumerate([links_1, links_2]):
    for link in link_group:
        if link is not None:
            filename = link.split('/')[-1]
            image = requests.get(link)
            if image.status_code == 200:
                with open(f'{number + 1}/{filename}', 'wb') as img:
                    img.write(image.content)
                    print(f'картинка {filename} скачана')
                    img.close()
            else: 
                print('Возникли траблы!!!')