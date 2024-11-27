import os

import bs4
from selenium import webdriver


def driver_init(url):
    if os.name == "posix":
        driver = webdriver.Safari()
        driver.get(url)
        return driver
    else:
        driver = webdriver.Firefox()
        driver.get(url)
        return driver


def get_html(driver):
    html = driver.page_source
    soup = bs4.BeautifulSoup(html, "html.parser")
    return soup


def get_cleaned_elements_from_first_column(soup_object):
    ignore_list = ("None", "открытый", "закрытый")
    cleaned_list = []

    for i in range(1, 51):
        class_name = f'field_fixed_{i}'
        row = soup_object.find('tr', class_=class_name)
        if row:
            field_name_element = row.find('td', class_='field_name')
            if field_name_element:
                link_element = field_name_element.find('a')
                if link_element and link_element.text not in ignore_list:
                    clean_el = " ".join(link_element.text.split()).strip()
                    if "закрытый" in clean_el:
                        clean_el = clean_el.replace("закрытый", "").strip()
                    elif "открытый" in clean_el:
                        clean_el = clean_el.replace("открытый", "").strip()
                    cleaned_list.append(clean_el)
    return cleaned_list


def check_page_nums(soup_object):
    nums_of_pages = soup_object.findAll('a', class_='js_pagination item')
    return nums_of_pages[-1].text


def get_cleaned_elements_from_main_table(soup_object, classname):
    lst = []
    raw_html = soup_object.findAll("td", class_=classname)

    for raw_el in raw_html:
        el = raw_el.find(class_="js_td_width")

        if el:
            # Если внутри элемента есть изображение
            img = el.find("img")
            if img and img.has_attr("title"):
                lst.append(img["title"])
            else:
                clean_el = " ".join(el.text.split())
                lst.append(clean_el)

    return lst
