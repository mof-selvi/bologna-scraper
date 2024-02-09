#%%

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from urllib.parse import urlparse, urljoin
import os
import time
import re
import xlsxwriter


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
INITIALIZED = False


options_ = Options()
options_.add_argument("--headless")  # Run Chrome in headless mode
options_.add_argument(f"user-agent={USER_AGENT}")

# service = Service(executable_path=r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')
service = Service(executable_path=r'chromedriver.exe')

# Use ChromeDriver to run Chrome in headless mode
driver = webdriver.Chrome(service=service, options=options_)

def links_read():
    all_lines=[]
    if os.path.isfile("links.txt"):
        with open("links.txt", "r") as f:
            while True:
                line = f.readline()
                if not line:
                    break
                all_lines.append(line.strip())
    return all_lines
    
def links_append(link):
    with open("links.txt", "a") as f:
        f.write(link+"\n")

def fetch(url, selector):
    global INITIALIZED

    if(INITIALIZED==False):
        driver.get(url)
        time.sleep(2)
        INITIALIZED = True
    
    rt = None
    
    try:
        if driver.current_url != url:
            driver.get(url)
            time.sleep(1.5)
        page_source = driver.page_source


        data = driver.find_elements(By.CSS_SELECTOR, selector)
        rt = data
        # for d in data:

        #     thedata = d.get_attribute('innerText')
        #     thedata_nospace = thedata.strip() #.replace("\t","").replace(" ","").replace("\r","").replace("\n","")
        #     if(thedata_nospace!=''):
        #         pass
    finally:
        pass

    return rt




# url_main = "https://obs.btu.edu.tr/oibs/bologna/progCourses.aspx?lang=tr&curSunit=6148"
url_main = "https://obs.btu.edu.tr/oibs/bologna/progCourses.aspx?lang=tr&curSunit=6237"



link_selector = "a:has(i.fa-info-square), a:has(i.fa-chevron-square-down)"

link_list = [] #links_read()
# if (len(link_list)>0):
#     print(link_list)
#     print("Let's continue...")
clicked_links = []

try:
    target_group_link = ""
    previous_groups = []

    while True:
        elems = fetch(url_main, link_selector)

        change_exists = False

        for elem in elems:
            href_value = elem.get_attribute('href')
            link_text = elem.get_attribute('innerText')

            if "grdBolognaDersler" not in href_value:
                continue

            # if 'Gruplu Dersleri Göster' in link_text:
            if 'DersAyrinti' not in href_value:
                print("Got an href for grouping:",href_value)
                # if target_group_link=="":
                #     target_group_link = href_value
                if target_group_link!=href_value and href_value not in previous_groups:
                    previous_groups.append(target_group_link)
                    target_group_link = href_value
                
                if target_group_link == href_value:
                    # clicked_links.append(href_value)
                    elem.click()
                    time.sleep(1.5)

                    change_exists = True
                    break
                else:
                    continue
            
            # if 'Gruplu Dersleri Gizle' in link_text:
            #     continue

            if href_value not in clicked_links:
                print("# ",link_text," >> ",href_value)
                clicked_links.append(href_value)
                elem.click()
                time.sleep(1.5)
                if driver.current_url not in link_list:
                    link_list.append(driver.current_url)
                    links_append(driver.current_url)
                print("Got the link:",driver.current_url)
                # driver.execute_script("window.history.go(-1)")
                # fetch(url_main, link_selector)
                change_exists = True
                break
        
        if change_exists == False:
            break
        
    print(link_list)
    print(len(link_list),"links have been collected.")


    
    print("The links have been appended to the file.")

finally:
    driver.quit()

print("All done.")