# Get-Ceiba-Scores

自動從[台大課程網站CEIBA](https://ceiba.ntu.edu.tw/index.php)以TA帳號登入，
將指定課程中的全部作業成績抓下來，並整理成一份excel


## 執行需求
 * [Python3](https://www.python.org/): 編寫版本為version 3.7
 * [Selenium](https://www.selenium.dev/): 模擬使用者的自動化操作套件，操作使用的的webdriver為chrome
 * [Beautifulsoup](https://www.crummy.com/software/BeautifulSoup/bs4/doc/): 解析及取得HTML原始碼各個標籤的元素資料
 * [time](https://docs.python.org/3/library/time.html): 防止短時間內大量存取而被Ban
 * [openpyxl](https://openpyxl.readthedocs.io/en/stable/): 將資料寫入excel


## 輸入資料
 * CEIBA account
 * CEIBA password
 * Assign semester: 如果要抓取當前學期的課程成績則直接按Enter跳過。但若要抓以前的課程(非本學期)，則需輸入指定學期(ex: 109-1)
