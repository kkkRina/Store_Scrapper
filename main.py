from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import csv

# Настройки драйвера
options = webdriver.ChromeOptions()
options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

url = "https://store.creality.com/collections/scanners"
driver.get(url)

try:
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product-item")))
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".product-item a")))

except Exception as e:
    print(f"Не удалось загрузить товары на странице: {e}")

driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(3)

links = set()  # используем set для исключения дубликатов

product_elements = driver.find_elements(By.CSS_SELECTOR, ".product-item a")

for product in product_elements:
    try:
        link = product.get_attribute("href")
        if link:
            links.add(link)
    except Exception as e:
        print(f"Ошибка при получении ссылки: {e}")
        continue

driver.quit()

# Сбор данных о товарах
options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

products_data = []

for url in links:
    driver.get(url)

    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product-main h1")))
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".product-price .price")))
    except Exception as e:
        print(f"Не удалось загрузить данные о товаре по ссылке {url}: {e}")
        continue

    try:
        title = driver.find_element(By.CSS_SELECTOR, ".product-main h1").text
    except:
        title = "Не найдено"

    try:
        price = driver.find_element(By.CSS_SELECTOR, ".product-price .price").text
    except:
        price = "Не указана"

    try:
        old_price = driver.find_element(By.CSS_SELECTOR, ".product-price .compareAtPrice").text
    except:
        old_price = "Не указана"


    try:
        product_info_item = driver.find_elements(By.CSS_SELECTOR, ".product-info-item")[1]
        shipping = product_info_item.find_element(By.CSS_SELECTOR, ".product-info-item-content span").text
    except Exception as e:
        shipping = "Не указано"
        print(f"Ошибка при извлечении срока доставки: {e}")

    # Добавляем данные о товаре в общий список
    products_data.append([title, price, old_price, shipping])

# Сохраняем все данные в CSV
with open("products_info.csv", "w", newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    writer.writerow(["Название", "Цена", "Старая цена", "Сроки доставки"])
    for product in products_data:
        writer.writerow(product)

print("Данные успешно сохранены в products_info.csv")

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Товары"

ws.append(["Название", "Цена", "Старая цена", "Срок доставки"])

for product in products_data:
    ws.append(product)

# Сохраняем все данные в xlsx
wb.save("products_info.xlsx")

print("Данные успешно сохранены в products_info.xlsx")

driver.quit()











