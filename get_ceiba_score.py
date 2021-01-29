from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import time
import openpyxl


url = "https://ceiba.ntu.edu.tw/index.php"
TA_account = input('Please input CEIBA account: ')
TA_password = input('Please input CEIBA password: ')
Semester = input('Assign semester(if u don\'t need to change, just push enter): ')

options = Options()
options.add_argument("--disable-notifications")
 
chrome = webdriver.Chrome(options = options)
chrome.get(url)

# enter ceiba
def choose_semester():
    if Semester == "":
        return
    
    chrome.find_element_by_css_selector("select[name='select_d0']").click()
    chrome.find_element_by_css_selector("option[value='index.php?seme_op=" + Semester + "']").click()


def get_scores(n):
    chrome.find_element_by_link_text("此作業所有列表").click()
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    sub_scores = soup.find('table').find_all(['tbody', 'input'], { 'name': 'old_rank_choice[]' })
    for i in sub_scores:
        hw_scores[n - 2].append(i.attrs['value'])

    chrome.find_element_by_link_text("批改作業").click()

def get_students():
    chrome.find_element_by_link_text("此作業所有列表").click()
    soup = BeautifulSoup(chrome.page_source, 'lxml')
    sub_students = soup.find('table').find_all(['tbody', 'tr'])
    for i, student in enumerate(sub_students):        
        if(i < 2):
            continue;

        student_info = student.find_all(['td', 'span'])
        students_id.append(student_info[5].text)
        students_name.append(student_info[7].text)

    chrome.find_element_by_link_text("批改作業").click()


chrome.find_element_by_css_selector("input[type='radio'][name='class'][value='1']").click()

chrome.find_element_by_name("loginid").send_keys(TA_account)
chrome.find_element_by_name("password").send_keys(TA_password)
chrome.find_element_by_css_selector("input[type='submit'][value='登入']").click()

time.sleep(1)

choose_semester()

chrome.find_element_by_name("b1").click()
chrome.find_element_by_xpath("//input[@onclick=\"singleadm('hw')\"]").click()
chrome.find_element_by_link_text("批改作業").click()
soup = BeautifulSoup(chrome.page_source, 'lxml')
hw_table = soup.find(['div', 'table'], {'id': 'sect_cont'}).find_all(['tbody', 'tr'])

hw_titles = []
students_id = []
students_name = []
hw_scores = [[] for _ in range(len(hw_table) - 2)]

# get hw_titles
for i, n in enumerate(hw_table):
    if i < 2:
        continue;

    info_list = n.find_all('td')
    hw_titles.append(info_list[1].text)

for i in range(1, len(hw_table)):
    if i < 2:
        continue;

    hw_path = '/html/body/div[1]/div[3]/div[2]/div/table/tbody/tr[' + str(i) + ']/td[7]/input'
    print(hw_titles[i-2])
    chrome.find_element_by_xpath(hw_path).click()
    if i == 2:
        get_students()
        chrome.find_element_by_xpath(hw_path).click()
        time.sleep(1)

    get_scores(i)
    time.sleep(1)


# data management
hw_data = [[] for _ in range(len(hw_scores)+2)]
hw_data[0].append('StudentID')
hw_data[0].extend(students_id)
hw_data[1].append('Name')
hw_data[1].extend(students_name)

for i in range(len(hw_titles)):
    hw_data[i + 2].append(hw_titles[i])
    hw_data[i + 2].extend(hw_scores[i])

wb_data = [['' for _ in range(len(hw_data))] for _ in range(len(hw_data[0]))]

for i in range(len(hw_data)):
    for j in range(len(hw_data[0])):
        wb_data[j][i] = hw_data[i][j]


# export to excel
wb = openpyxl.Workbook()
sheet = wb.create_sheet("離岸風力發電導論", 0)

for data in wb_data:
    sheet.append(data)

wb.save('成績.xlsx')
