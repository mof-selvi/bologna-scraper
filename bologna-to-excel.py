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


def shorten_name(text:str):
    text = re.sub(r'[^a-zçöışüğA-ZÇÖİŞÜĞ0-9_\s]+', '', text)
    text = text[:31]
    return text

def table2list(text:str):
    rt = []
    text = text.replace("\r","")
    table_lines = text.split("\n")
    for l in table_lines:
        rt.append(l.split("\t"))
    return rt

def substring_between(text:str, str1:str, str2:str):
    s=str(re.escape(str1))
    e=str(re.escape(str2))
    rt=re.findall(s+"(.*)"+e,text)[0]
    return rt


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36"
INITIALIZED = False



options_ = Options()
options_.add_argument("--headless")  # Run Chrome in headless mode
options_.add_argument(f"user-agent={USER_AGENT}")

# service = Service(executable_path=r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe')
service = Service(executable_path=r'chromedriver.exe')

# Use ChromeDriver to run Chrome in headless mode
driver = webdriver.Chrome(service=service, options=options_)



def download_bologna(url):
    global INITIALIZED

    if(INITIALIZED==False):
        driver.get(url)
        time.sleep(5)
        INITIALIZED = True

    try:
        driver.get(url)
        time.sleep(1.5)
        page_source = driver.page_source
        # print(page_source)
        # print("-"*30)

        urlid = substring_between(url,"curCourse=","&")
        workbook_name = urlid+"_NONAME"


        # Prepare the workbook name
        initial_data_tables = driver.find_elements(By.CSS_SELECTOR, "table")
        if(len(initial_data_tables)>1):
            initial_data_table = initial_data_tables[1]
            initial_data = initial_data_table.get_attribute('innerText')
            # print(initial_data)
            initial_data_nospace = initial_data.strip()
            if(initial_data_nospace!=''):
                initial_data_list = table2list(initial_data)
                if(len(initial_data_list)>1 and len(initial_data_list[1])>2):
                    workbook_name = urlid+'_'+initial_data_list[1][1] + ' ' + initial_data_list[1][2]
                else:
                    print("#"*10)
                    print("Error: URL has no data. @",urlid)
                    print(url)
                    print("#"*10)
                    return

        print("Workbook name:")
        print(workbook_name)


        # Prepare the Excel file
        workbook = xlsxwriter.Workbook("downloads/"+workbook_name+".xlsx")
        worksheet = None
        # worksheet = workbook.add_worksheet()
        # Prepare formats
        format_title = workbook.add_format({"bold": True})
        format_table = workbook.add_format({'border': 1, 'text_wrap': True})

        document = []



        try:
            data = driver.find_elements(By.CSS_SELECTOR, "table:has(div)")
            for d in data:
                oldhtml = d.get_attribute('innerHTML')
                newhtml = oldhtml.replace('<div','<nodiv').replace('</div','</nodiv')
                driver.execute_script("arguments[0].innerHTML=arguments[1];",d, newhtml)


            data = driver.find_elements(By.CSS_SELECTOR, "td:has(br)")
            for d in data:
                oldhtml = d.get_attribute('innerHTML')
                newhtml = oldhtml.replace('<br>',', ').replace('<br ',', ')
                driver.execute_script("arguments[0].innerHTML=arguments[1];",d, newhtml)
        except:
            pass



        # fetch all data
        # tables = driver.find_elements(By.TAG_NAME, 'table')
        # data = driver.find_elements(By.CSS_SELECTOR, "table, span:not(table span)")
        data = driver.find_elements(By.CSS_SELECTOR, ".panel table, .panel .panel-heading > span")
        for d in data:

            thedata = d.get_attribute('innerText')
            thedata_nospace = thedata.strip() #.replace("\t","").replace(" ","").replace("\r","").replace("\n","")
            if(thedata_nospace!=''):

                if(d.tag_name=='span'):
                    # print(">",thedata)

                    document.append([shorten_name(thedata),thedata,[]])
                    
                    # worksheet = workbook.add_worksheet(shorten_name(thedata))
                    # doc_line = 0
                    # worksheet.write(0,0,thedata,format_title)
                    # worksheet.autofit()


                elif(d.tag_name=='table'):

                    # if(worksheet == None):
                    #     worksheet = workbook.add_worksheet()

                    thelist = table2list(thedata)
                    document[-1][2] = thelist

                    # print(thelist)
                    # print("#"*20,"\n\n\n")
                    # for therow in thelist:
                    #     worksheet.write_row(2,1,therow,format_table)
                    # worksheet.autofit()

                # print(d.tag_name)
                # print("???")
                # print(thedata)
                # print("#"*20)
        

        # print(document)
        
        for d in document:
            if(len(d[2])>0):
                worksheet = workbook.add_worksheet(d[0])
                worksheet.write(0,0,d[1],format_title)
                rowidx=2

                col_widths = {}

                for therow in d[2]:
                    worksheet.write_row(rowidx,1,therow,format_table)
                    rowidx+=1

                    colidx = 0
                    for thecol in therow:
                        colidxstr = str(colidx)
                        thecollen = len(thecol)+5
                        if colidxstr not in col_widths.keys():
                            col_widths[colidxstr] = thecollen
                        else:
                            if(col_widths[colidxstr]<thecollen):
                                col_widths[colidxstr] = thecollen
                        colidx+=1
                    
                
                for c,w in col_widths.items():
                    worksheet.set_column(1+int(c),1+int(c),min(int(w),100))
                # worksheet.autofit()

        # Close the Excel file
        workbook.close()

    finally:
        pass


    print("Done! File has been saved in the current working directory.")


# url="https://obs.btu.edu.tr/oibs/bologna/progCourseDetails.aspx?curCourse=251941&lang=tr"
# url="https://obs.btu.edu.tr/oibs/bologna/progCourseDetails.aspx?curCourse=251931&lang=tr"
# url="https://obs.btu.edu.tr/oibs/bologna/progCourseDetails.aspx?curCourse=251961&lang=tr"

try:
    f = open("links.txt", "r")
    url_lines = f.readlines()
    f.close()

    for url in url_lines:
        url = url.strip()
        if(url!=''):
            download_bologna(url)
finally:
    driver.quit()

print("All done.")

