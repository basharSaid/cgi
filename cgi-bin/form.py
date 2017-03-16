#!/usr/bin/env python3
import cgi
import html
import xlrd
import cssutils
import cgitb; cgitb.enable()

# Задаем переменные
form = cgi.FieldStorage()
a = form.getfirst("TEXT_1", "не задано")
a = html.escape(a)
a = a.title()

print("Content-Type: text/html; charset=utf-8\r")
print('\r')
print("""<!DOCTYPE HTML>
<html>

<head>
  <meta charset="utf-8">
  <link rel="stylesheet" href="../css/main.css" type="text/css" charset="utf-8">
    <title>Телефоны</title>
</head>

    <header>
  			<img src="../img/logo.png" width="5%" height="5%" title="TEST" alt="test">
    </header>

      <form action="/cgi-bin/form.py">
          <input type="text" name="TEXT_1">
          <input type="submit">
      </form>
""")

print("<br>")
print("<h4>Запрос обработан</h4>")

# Код 
rb = xlrd.open_workbook(r'./Список телефонов.xls')
sheet = rb.sheet_by_index(0)
vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
for kabinet, otdel, gorod, phone, dolgnost, first_name in vals:
    first_name = first_name.strip()
    if a == first_name[:1] or a == first_name[:2] or a == first_name[:3] \
            or a == first_name[:4] or a == first_name[:4] or a == first_name[:5] \
            or a == first_name[:6] or a == first_name[:7] or a == first_name[:8]:
        kabinet = str(kabinet)
        kabinet = kabinet[:3]
        phone = str(phone)
        phone = phone[:3]
        name = first_name, kabinet, otdel, gorod, phone, dolgnost
        print('{} : каб.{} : {} : {} : т. {} : {}'.format(*name) + '<br>')
    elif a != first_name:
        pass

# Закрываем html
print("""</body>
        </html>""")