import requests  # pip install requests
from bs4 import BeautifulSoup  # pip install bs4
from openpyxl import Workbook
import os
import random
import json
import time

# pip install lxml

# Список пользовательских агентов
user_agents = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Gecko/20100101 Firefox/113.0',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Gecko/20100101 Firefox/112.0',
    'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
    'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1',
    'Mozilla/5.0 (Linux; Android 11; SM-G998B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.77 Mobile Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:78.0) Gecko/20100101 Firefox/78.0',
    'Mozilla/5.0 (Linux; Android 10; SM-G960F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Mobile Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Safari/605.1.15',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 OPR/76.0.4017.177',
    'Mozilla/5.0 (Linux; Android 11; Pixel 4 XL Build/RQ3A.210705.001) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Mobile Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:86.0) Gecko/20100101 Firefox/86.0',
    'Mozilla/5.0 (Linux; Android 5.1; Nexus 5 Build/LMY48B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Mobile Safari/537.36'
]


# Функция для получения случайного пользовательского агента
def get_random_user_agent():
    return random.choice(user_agents)


# Создание нового XLSX файла
workbook = Workbook()
sheet = workbook.active

print('Ведите марку изделия, например: аввг...')
mark_num = input()
# print('Введите кол-во страниц...')
# page_num = input()
sota = []
count = 1
while count <= 28:
    url = f'https://www.etm.ru/catalog?page=28&rows=36&searchValue={mark_num}'
    headers = {
        'User-Agent': get_random_user_agent(),
        'Accept': '*/*',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive'  # Исправлено здесь
    }
    data = requests.get(url, headers=headers).text
    block = BeautifulSoup(data, 'lxml')
    heads = block.find_all('a', {
        'class': 'MuiTypography-root MuiTypography-inherit MuiLink-root MuiLink-underlineHover tss-lrg5ji-root-blue-title mui-style-i8aqv9'})
    print(len(heads))
    if not heads:
        print("Больше страниц не найдено.")
    for head in heads:
        link = head.get('href')
        print('https://www.etm.ru' + link)
        get_link = ('https://www.etm.ru' + link)
        spot = requests.get(get_link, headers=headers).text
        time.sleep(2)
        try:
            soup = BeautifulSoup(spot, 'lxml')
            category = soup.find('ul', {'class': 'tss-52eym-root'}).find_all('li')
            # Создаём список для хранения значений
            values = []

            # Итерация по категориям
            for i in category:
                # Добавляем значение в список
                values.append(i.text.strip().replace('/', ' '))

            # Проверяем, что в списке достаточно значений
            if len(values) >= 3:
                # Выводим 3 последних значения по отдельности
                print(values[-3])
                get_category = (values[-3])
                print(values[-2])
                podkategory = (values[-2])
                print(values[-1])
                last_kategory = (values[-1])
            else:
                print("Недостаточно значений для вывода.")
            name = ' '.join(soup.find('h1').text.strip().split())
            print(name)
            try:
                price = soup.find('span', {'class': 'tss-1de1yzy-priceMain'}).text.strip()
                print(price)
            except:
                print('None')
                price = 'None'
            articul = soup.find('p', {
                'class': 'MuiTypography-root MuiTypography-body2 tss-5nwafn-text mui-style-1xmqc91'}).text.strip()
            print(articul)

            # Ищем все div с классом 'tss-cz0zic-zoomIcon'
            photo_elements = soup.find_all('div', {'class': 'tss-cz0zic-zoomIcon'})

            # Проходим по найденным элементам и выводим src изображений с нужными расширениями
            get_photo = []
            for photo in photo_elements:
                img_tag = photo.find('img')
                if img_tag and 'src' in img_tag.attrs:
                    img_src = img_tag['src']
                    # Проверяем, заканчивается ли источник на .jpg или .png
                    if img_src.endswith('.jpg') or img_src.endswith('.png'):
                        # print(img_src)
                        get_photo.append(img_src)

            params = soup.find('div',
                               {
                                   'class': 'MuiGrid-root MuiGrid-container MuiGrid-spacing-lg-4 mui-style-1sc115b'}).find_all(
                'tr')
            # Ваш существующий код
            totals = []  # Список для сохранения строк total

            for param in params:
                all_params = param.find_all_next('td')
                param_key = all_params[0].text.strip()
                param_value = all_params[1].text.strip()
                print(param_key)
                print(param_value)
                total = param_key + ': ' + param_value
                totals.append(total)  # Добавляем строку total в список
                print(total)
            # Подготовим словарь с параметрами
            param_dict = {}
            for i, value in enumerate(totals, start=1):
                param_dict[f'param_{i}'] = value  # Генерация ключей param_1, param_2 и т.д.

            print('\n')
            discription = soup.find('div', {'class': 'tss-f73qk4-content'}).find('p').text.strip()
            print(discription)
            print('\n')
            print('\n')
            discription = soup.find('div', {'class': 'tss-f73qk4-content'}).find('p').text.strip()
            print(discription)
            print('\n')

            # Предположим, что у вас уже есть заполненный словарь storage
            # Ваш существующий словарь storage
            storage = {
                'category': get_category,
                'podkategory': podkategory,
                'last_kategory': last_kategory,
                'name': name,
                'price': price,
                'discription': discription,
                **param_dict  # Объединение словарей
            }

            sota.append(storage)

            # Запись словаря в файл в формате JSON
            # with open(f'{tema}.json', 'w', encoding='utf-8') as json_file:
            #     json.dump(sota, json_file, ensure_ascii=False, indent=4)
            # directory = f'C:\\Users\\user\\PycharmProjects\\etm_cabely_pars\\{last_kategory}'
            # if not os.path.exists(directory):
            #     os.makedirs(directory)

            with open(f'{mark_num}.json', 'w', encoding='utf-8') as json_file:
                json.dump(sota, json_file, ensure_ascii=False, indent=4)

        except:
            print('None')
            continue
    count += 1
    print('Page_num: ' + str(count))
