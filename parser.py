import bs4
import lxml
from selenium import webdriver


def driver_init(url):
    driver = webdriver.Firefox()
    driver.get(url)
    return driver


def get_html(driver):
    html = driver.page_source
    soup = bs4.BeautifulSoup(html, "lxml")
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


def get_cleaned_elements_from_main_table(soup_object, classname):
    lst = []
    raw_html = soup_object.findAll("td", class_=classname)
    for raw_el in raw_html:
        el = raw_el.find(class_="js_td_width")
        if el is None:
            el = "-"
            lst.append(el)
        else:
            clean_el = " ".join(el.text.split())
            lst.append(clean_el)
    return lst


if __name__ == "__main__":
    pass
