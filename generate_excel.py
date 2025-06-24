import os
import time
import re
from datetime import datetime
from bs4 import BeautifulSoup
import pdfkit
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import requests
import xlsxwriter

# 配置 wkhtmltopdf 路径（如有需要可修改）
config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf')

base_url = "https://medium.com/@sohail_saifi/"
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)

def safe_filename(s):
    return re.sub(r'[^\u4e00-\u9fa5\w\- ]', '', s).strip().replace(' ', '_')

# 启动 Selenium 浏览器
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # 无头模式
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
options.add_argument('--disable-blink-features=AutomationControlled')
print('Start browser（启动浏览器）...')
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
print('Open page（打开页面）...')
driver.get(base_url)
print('Start scrollin（开始滚动）...')
# 自动滚动页面，最多滚动3次，便于测试
# max_scroll_times = 3
scroll_times = 0
last_height = driver.execute_script("return document.body.scrollHeight")
while True:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)
    new_height = driver.execute_script("return document.body.scrollHeight")
    # scroll_times += 1
    if new_height == last_height: # or scroll_times >= max_scroll_times:
        break
    last_height = new_height
print('Scrolling ends, start parsing page（滚动结束，开始解析页面）...')
with open('page_source.html', 'w', encoding='utf-8') as f:
    f.write(driver.page_source)
print('Page source code has been saved to（页面源码已保存到） page_source.html')

# 解析 page_source.html
with open('page_source.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

divs = soup.find_all('div', attrs={'role': 'link', 'data-href': True})
print('Number of articles found（找到的文章数量）:', len(divs))

# 用于存储表格数据
article_rows = []

for i, div in enumerate(divs):
    article_url = div['data-href']
    print(f"Generated table（解析文章链接）: {article_url}")
    try:
        art_resp = requests.get(article_url)
        art_soup = BeautifulSoup(art_resp.text, 'html.parser')
        title_tag = art_soup.find('h1')
        title = title_tag.text.strip() if title_tag else f"article_{i+1}"
        # 查找 span[data-testid=storyPublishDate] 并解析英文日期
        date_str = None
        story_publish = art_soup.find('span', attrs={"data-testid": "storyPublishDate"})
        if story_publish:
            date_text = story_publish.get_text(strip=True)
            try:
                pub_date = datetime.strptime(date_text, '%b %d, %Y')
                date_str = pub_date.strftime('%Y-%m-%d')
            except Exception:
                date_str = date_text  # 解析失败就用原文本
        if not date_str:
            date_str = datetime.now().strftime('%Y-%m-%d')
        filename = f"{date_str}_{safe_filename(title)}.pdf"
        article_rows.append([i+1, title, date_str, filename, article_url])
    except Exception as e:
        print(f"解析 {article_url} 时出错: {str(e)}")

# 生成xlsx表格
excel_file = os.path.join(output_dir, 'articles.xlsx')
workbook = xlsxwriter.Workbook(excel_file)
worksheet = workbook.add_worksheet('Articles')
# 写表头
headers = ['Index', 'Title', 'Date', 'Filename', 'Link']
for col, h in enumerate(headers):
    worksheet.write(0, col, h)
# 写数据
for row_idx, row in enumerate(article_rows, start=1):
    for col_idx, value in enumerate(row):
        worksheet.write(row_idx, col_idx, value)
workbook.close()
print(f"已生成表格: {excel_file}")

driver.quit()  # 最后关闭