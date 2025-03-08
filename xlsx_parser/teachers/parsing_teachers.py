import requests
from bs4 import BeautifulSoup
import time
teachers = []

base_url = "https://www.vstu.ru/university/personalii/?searchstring="
response = requests.get(base_url)
soup = BeautifulSoup(response.text, features="html.parser")
all_names = soup.find_all("dd", {"class": "name"})
for name in all_names:
    teachers.append(name.text)


for page in range(2, 36):
    response = requests.get(base_url + f"&PAGEN_1={page}")
    soup = BeautifulSoup(response.text, features="html.parser")
    all_names = soup.find_all("dd", {"class": "name"})
    for name in all_names:
        teachers.append(name.text)

    time.sleep(1)


import json
with open('teachers.json', 'a', encoding='utf-8') as file:
    json.dump(teachers, file, ensure_ascii=False, indent=4)


