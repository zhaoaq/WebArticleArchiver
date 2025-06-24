import os
import time
import re
import random
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from PIL import Image
import img2pdf

base_url = "https://medium.com/@sohail_saifi/"
output_dir = 'output'
os.makedirs(output_dir, exist_ok=True)


def safe_filename(s):
    """生成安全的文件名"""
    return re.sub(r'[^\u4e00-\u9fa5\w\- ]', '', s).strip().replace(' ', '_')


def full_page_screenshot(driver, filename):
    """
    截取整个页面的长截图
    """
    # 获取原始窗口大小
    original_size = driver.get_window_size()

    # 获取页面总高度
    total_height = driver.execute_script("return document.body.parentNode.scrollHeight")

    # 设置窗口大小为页面总高度
    driver.set_window_size(1200, total_height)

    # 等待页面调整
    time.sleep(1)

    # 截取整个页面
    driver.save_screenshot(filename)

    # 恢复原始窗口大小
    driver.set_window_size(original_size['width'], original_size['height'])

    return filename


def download_pdfs(skip=0, limit=0, max_retries=3):
    """下载PDF文件（支持跳过和限制数量）"""
    # 读取Excel文件
    excel_file = os.path.join(output_dir, 'articles.xlsx')
    if not os.path.exists(excel_file):
        print(f"错误: 文件 {excel_file} 不存在")
        return

    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active
    articles = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row:  # 确保行不为空
            articles.append(row)

    total = len(articles)
    print(f"Excel中找到 {total} 篇文章")

    # 应用跳过和限制
    if skip > 0 or limit > 0:
        start_idx = skip
        end_idx = skip + limit if limit > 0 and (skip + limit) <= total else total
        articles = articles[start_idx:end_idx]
        print(f"实际下载范围: 第 {start_idx + 1} 到 {end_idx} 篇")

    # 配置浏览器选项
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')  # 解决内存问题
    options.add_argument(
        'user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36')
    options.add_argument('--disable-blink-features=AutomationControlled')

    # 启动浏览器
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # 设置页面加载超时
    driver.set_page_load_timeout(30)

    try:
        # 下载PDF
        for idx, article in enumerate(articles):
            article_idx, title, date_str, filename, article_url = article
            pdf_file = os.path.join(output_dir, filename)

            # 检查标题是否包含"Weekly Report"（不区分大小写）
            if "weekly report" in title.lower():
                print(f"[{idx + 1}/{len(articles)}] 跳过: {title} (包含'Weekly Report')")
                continue

            print(f"[{idx + 1}/{len(articles)}] 处理: {title[:50]}...")

            # 检查标题是否包含"SR RANKING"（不区分大小写）
            # if "SR RANKING" in title.lower():
            #     print(f"[{idx + 1}/{len(articles)}] 跳过: {title} (包含'SR RANKING')")
            #     continue
            #
            # print(f"[{idx + 1}/{len(articles)}] 处理: {title[:50]}...")

            # 重试机制
            retry_count = 0
            success = False

            while retry_count < max_retries and not success:
                try:
                    # 添加随机延迟防止被封 (2-4秒)
                    delay = 2 + random.random() * 2
                    time.sleep(delay)

                    # 访问页面
                    print(f"  访问页面: {article_url}")
                    driver.get(article_url)

                    # 等待主要内容加载
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.TAG_NAME, 'article'))
                    )

                    # 添加额外等待确保页面完全加载
                    time.sleep(1.5)

                    # 生成临时截图文件名
                    temp_image = os.path.join(output_dir, f"temp_{article_idx}.png")

                    # 截取完整页面
                    print("  截取页面截图...")
                    full_page_screenshot(driver, temp_image)

                    # 将截图转换为PDF
                    print("  将截图转换为PDF...")
                    with open(temp_image, "rb") as f:
                        image = Image.open(f)
                        # 转换为RGB模式（如果原始是RGBA）
                        if image.mode == 'RGBA':
                            image = image.convert('RGB')

                        # 保存为PDF
                        image.save(pdf_file, "PDF", resolution=100.0)

                    # 删除临时截图
                    os.remove(temp_image)

                    print(f"√ 已保存: {pdf_file}")
                    success = True

                except Exception as e:
                    retry_count += 1
                    print(f"× 尝试 {retry_count}/{max_retries} 失败 (文章 {article_idx}): {str(e)}")

                    # 捕获特定错误类型
                    if "timeout" in str(e).lower():
                        print("  页面加载超时，尝试刷新页面...")

                    if retry_count < max_retries:
                        # 等待一段时间后重试（增加等待时间）
                        wait_time = 5 + random.random() * 5
                        print(f"  等待 {wait_time:.1f} 秒后重试...")
                        time.sleep(wait_time)
                    else:
                        print(f"× 下载失败 (文章 {article_idx}): 达到最大重试次数")
                        print(f"程序将在下载第 {article_idx} 篇文章后退出")
                        return

        print("所有文章下载完成！")

    finally:
        driver.quit()
        print('浏览器已关闭')


if __name__ == "__main__":
    # 第二步：下载PDF（可多次执行）
    # 参数说明：
    # skip - 跳过的文章数量（从0开始计数）
    # limit - 限制下载的文章数量（0表示无限制）
    # max_retries - 每篇文章最大重试次数（默认为3）

    # 测试少量文章
    download_pdfs(skip = 2, limit = 3, max_retries=3)