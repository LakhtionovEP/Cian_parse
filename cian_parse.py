import urllib.request
import datetime
import re
import openpyxl
from xml.etree import ElementTree as ET
from bs4 import BeautifulSoup
import requests
import os

# получение данных с cbr.ru по ссылке


def get_obl_cb(url):
    a = []
    response = requests.get(url)
    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
    obj = soup.find_all('table', class_="data")
    for line in str(obj).split('\n'):
        a.append(line)
    stavka_cb = a[-3].lstrip('<td>').rstrip('</td>')
    stavka_cb = float(re.sub(',', '.', stavka_cb))
    return stavka_cb

# получение данных с cian по ссылке


def get_cian_data(url):

    # использование регулярных выражений для очистки данных от шума

    filr = re.compile(r'[\"]>(.*?)</')
    filra = re.compile(r'[\"]>([0-9а-яА-Я.\W]*?)</')
    filrd = re.compile(r'>\s+(\d.*?)</')
    filrfl = re.compile(r'\d+\s[из]+\s\d+')
    fillrtype = re.compile(r'(\b\Монолитный\b)|(\b\Кирпичный\b)|(\b\Панельный\b)')
    fillmt = re.compile(r'\d+')
    response = requests.get(url)
    html = response.content
    soup = BeautifulSoup(html, 'html.parser')

    # парсинг html с cian.ru

    obj = soup.findAll('span', attrs={'class': "a10a3f92e9--underground_time--1fKft"})
    obj2 = soup.findAll('a', attrs={'class': "a10a3f92e9--underground_link--AzxRC"})
    obj3 = soup.find('span', attrs={'itemprop': "price"})
    obj4 = soup.findAll('a', attrs={'class': "a10a3f92e9--phone--3XYRR"})
    obj5 = soup.find('div', attrs={'class': "a10a3f92e9--price_per_meter--hKPtN a10a3f92e9--price_per_meter--residential--1mFDW"})
    obj6 = soup.findAll('address', attrs={'class': "a10a3f92e9--address--140Ec"})
    obj7 = soup.find('div', attrs={'class': "a10a3f92e9--info-value--18c8R"})
    obj8 = soup.find('h1', attrs={'class': "a10a3f92e9--title--2Widg"})
    obj9 = soup.find_all('div', attrs={'class': "a10a3f92e9--info-value--18c8R"})
    obj10 = soup.find_all('div', attrs={'class': "a10a3f92e9--offer_card_page-bti--2BrZ7"})

    phone_str, add_str, metro_str  = '', '', ''
    a = []
    metro_time = filrd.findall(str(obj))  # время до метро
    if len(metro_time) == 0:
        metro_time = 'N/A'
    try:
        metro_name = filr.findall(str(obj2))  # название метро
    except:
        metro_name = 'N/A'
    try:
        flat_price = filr.findall(str(obj3))[-1]  # цена квартиры
    except:
        flat_price = 'N/A'
    try:
        phone = filr.findall(str(obj4))  # телефон(-ы)
        phone_str = ',\n'.join(phone)
    except:
        phone_str = 'N/A'
    try:
        flat_price_pm = filr.findall(str(obj5))[-1]  # цена квадратного метра
    except:
        flat_price_pm = 'N/A'
    try:
        address = filra.findall(str(obj6))  # адресс
        if address[-1] == 'На карте':
            address.pop(-1)
        add_str = ', '.join(address)
    except:
        add_str = 'N/A'
    try:
        square = filr.findall(str(obj7))[-1]  # площадь
    except:
        square = 'N/A'
    try:
        room_q = filr.findall(str(obj8))[-1].split(',')[0]  # кол-во комнат
    except:
        room_q = 'N/A'
    try:
        cur_fl, ov_fl = filrfl.findall(str(obj9))[-1].split(' из ')  # этаж и этажность здания
    except:
        cur_fl, ov_fl = 'N/A', 'N/A'
    try:
        b_type = fillrtype.search(str(obj10))[0]  # тип здания
    except:
        b_type = 'N/A'
    try:
        for i in range(len(metro_name)):  # выбор наименьшего пути до метро, путь на транспорте считается с коэф. = 4
            if 'пешком' in str(metro_time[i]):
                a.append(int(fillmt.findall(metro_time[i])[0]))
            else:
                a.append(int(fillmt.findall(metro_time[i])[0])*4)
        metro_str = min(a)
    except:
        metro_str = 'N/A'
    return metro_str, flat_price, phone_str, flat_price_pm, add_str, square, room_q, cur_fl, ov_fl, b_type

# преобразование и запись данных в Excel


def excel_output(url, column):
    metro, flat_price, phone, \
    flat_price_pm, address, square, room_q, cur_fl, ov_fl, b_type = get_cian_data(url)
    if int(room_q[0]) == 1:
        room_q_n = 'Однокомнатная'
        room_q_type = '1 комн.'
    elif int(room_q[0]) == 2:
        room_q_n = 'Двухкомнатная'
        room_q_type = '2 комн.'
    elif int(room_q[0]) >= 3:
        room_q_n = 'Многокомнатная'
        room_q_type = '3 комн. и более'
    else:
        room_q_n = '?'
        room_q_type = '?'

    if b_type == 'Кирпичный':
        b_type_t = 'кирпичном'
        wall_type = 'Кирпичные стены'
    elif b_type == 'Монолитный':
        b_type_t = 'монолитном'
        wall_type = 'Монолитные стены'
    elif b_type == 'Панельный':
        b_type_t = 'панельном/блочном'
        wall_type = 'Панельные/блочные стены'
    else:
        b_type_t = '?'
        wall_type = '?'
    sheet[str(column)+'7'] = f'{room_q_n} квартира в {b_type_t} жилом доме. Хорошее местоположение, удобные подъездные пути'
    sheet[str(column)+'8'] = f'Информационная база «ЦИАН», www.cian.ru, \n т. {phone}'
    sheet[str(column)+'9'] = url
    sheet[str(column)+'10'] = int(flat_price.replace(" ", "")[:-1])
    sheet[str(column)+'16'] = address

    if metro.isdigit():  # Проверка на наличие данных (в городе может не быть метро) + разбивка по интервалам
        if metro <= 5:
            metro_time_dur = 'До 5 минут пешком'
        elif 5 < metro <= 15:
            metro_time_dur = '5-15 минут пешком'
        elif 15 < metro <= 30:
            metro_time_dur = '15-30 минут пешком'
        elif 30 < metro <= 60:
            metro_time_dur = '30-60 минут пешком'
        elif 60 < metro <= 90:
            metro_time_dur = '60-90 минут пешком'
        else:
            metro_time_dur = '?'
    else:
        metro_time_dur = '?'

    sheet[str(column)+'17'] = metro_time_dur
    sheet[str(column)+'21'] = float(square.replace(",", ".")[:-2])
    sheet[str(column)+'23'] = wall_type
    sheet[str(column)+'25'] = room_q_type
    sheet[str(column)+'26'] = cur_fl
    sheet[str(column)+'27'] = ov_fl

# Обработка i-го объекта


def url_input(i):
    print(f'Вставьте ссылку на объект №{i} в ЦИАНе')
    symb = chr(66 + int(i)*2)  # преобразование для определения нужного столбца для записи в Excel
    while True:
        try:
            url = input()
            requests.get(url)
            print('Объект найден')
            excel_output(url, symb)
            print(f'Данные по объекту №{i} внесены\n')
            break
        except:
            print('Ссылка некорректна')
            continue


print('Введите 1 для отображения данных на сегодня, либо дату для архива')
daychoice = input()

print('Какой формат аналогов? (4 или 5)')  # выбор формата и файла для записи
while True:
    ana_qua = input()
    if ana_qua == '4':
        f_name = 'Шаблон-4.xlsx'
        break
    elif ana_qua == '5':
        f_name = 'Шаблон-5.xlsx'
        break
    else:
        print('Вы ввели некорректное значение')
        f_name = ''
        continue

wb = openpyxl.open(f_name)
sheet = wb.active
print('Идет проверка, что результирующий файл закрыт...')
while True:
    try:
        wb.save(f_name)
        break
    except PermissionError:
        print('Сначала нужно закрыть результирующий файл')
        os.system('pause')
        continue

while True:  # пользовательский выбор по необходимости внесения объектов
    print('Данные по какому объекту нужно внести? (0 - чтобы внести все)')
    obj_choice = input()
    if int(obj_choice) in range(int(ana_qua)+1):
        if obj_choice == '0':
            for s in range(int(ana_qua)):
                url_input(s+1)
            break
        else:
            url_input(obj_choice)
            print('Хотите внести данные по другому объекту (1 - да, любая другая - нет/выход)')
            obj_choice_n = input()
            if obj_choice_n == '1':
                continue
            else:
                print('Идет выход из программы, пожалуйста, подождите...')
                break
    else:
        print('Вы ввели некорректное значение')
        continue

# запись выбранной пользователем даты
if daychoice == '1':
    todayd = datetime.date.today().strftime('%d/%m/%Y')
else:
    todayd = daychoice
sheet['C1'] = todayd

# получение данных по курсу доллару на требуемый день
valuta = ET.parse(urllib.request.urlopen("http://www.cbr.ru/scripts/XML_daily.asp?date_req=" + str(todayd)))
id_dollar = "R01235"
for line in valuta.findall('Valute'):
    id_v = line.get('ID')
    if id_v == id_dollar:
        rub = float(re.sub(',','.',line.find('Value').text))
        sheet['C2'] = rub


# получение безрисковой ставки на требуемый день (данные часто запаздывают, поэтому используется бесконечный спуск)
while True:
    try:
        stavka_cb = get_obl_cb('https://cbr.ru/hd_base/zcyc_params/zcyc/?DateTo=' + str(todayd))
        sheet['C3'] = stavka_cb / 100
        break
    except:
        todayd = (datetime.datetime.strptime(todayd, "%d/%m/%Y").date()
                                  - datetime.timedelta(1)).strftime('%d/%m/%Y')
        continue

wb.save(f_name)
