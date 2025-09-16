from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd
import os
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.keys import Keys

# 配置参数
DRIVER_PATH = r"C:\Users\32597\Downloads\edgedriver_win64\msedgedriver.exe"  # 确保驱动路径正确
SEARCH_KEYWORD = "销售"  # 你可以修改为任意关键词


start_page = 399  # 你的起始页
end_page = 412

existing = set()
csv_path = "栏目详情.csv"
if os.path.exists(csv_path):
    df_exist = pd.read_csv(csv_path)
    for _, row in df_exist.iterrows():
        existing.add(f"{row['栏目名称']}_{row['页码']}")
else:
    df_exist = pd.DataFrame()

# 初始化浏览器
def init_browser():
    options = webdriver.EdgeOptions()
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors=yes")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--disable-web-security")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36 Edg/126.0.0.0")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    service = Service(DRIVER_PATH)
    browser = webdriver.Edge(service=service, options=options)
    browser.set_window_size(1200, 900)
    # 注入JS隐藏webdriver特征
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.navigator.chrome = { runtime: {} };
        Object.defineProperty(navigator, 'languages', {get: () => ['zh-CN', 'zh']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        """
    })
    return browser

# 选择地区（广东 - 江门 - 确定）
def select_region(browser):
    try:
        guangdong_btn = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[contains(text(),'广东省')]]"))
        )
        guangdong_btn.click()
        print("已点击广东省")
        jiangmen_btn = WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[contains(text(),'江门市')]]"))
        )
        jiangmen_btn.click()
        print("已点击江门市")
        WebDriverWait(browser, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[contains(text(),'确') and contains(text(),'定')]]"))
        )
        confirm_btn = browser.find_element(By.XPATH, "//button[span[contains(text(),'确') and contains(text(),'定')]]")
        ActionChains(browser).move_to_element(confirm_btn).pause(0.5).click(confirm_btn).perform()
        print("已点击确 定按钮")
        time.sleep(2)
        return True
    except Exception as e:
        print(f"地区选择失败: {e}")
        browser.save_screenshot("region_error.png")
        with open("region_page.html", "w", encoding="utf-8") as f:
            f.write(browser.page_source)
        return False

def close_close_popup_if_exists(browser, max_tries=5):
    try:
        tries = 0
        while tries < max_tries:
            close_btns = browser.find_elements(By.XPATH, "//button[span[contains(text(),'关') and contains(text(),'闭')]]")
            visible_btns = [btn for btn in close_btns if btn.is_displayed() and btn.is_enabled()]
            if not visible_btns:
                break
            for btn in visible_btns:
                try:
                    browser.execute_script("arguments[0].click();", btn)
                    print("已自动点击关 闭按钮")
                    time.sleep(0.5)
                except Exception as e:
                    print("点击关 闭失败，重试...", e)
                    continue
            tries += 1
            time.sleep(0.5)
        if tries == max_tries:
            print("警告：关 闭弹窗未能成功关闭，已强制跳出循环")
    except Exception as e:
        print("关闭关 闭弹窗异常：", e)

def extract_modal_info_selenium(browser, expected_title):
    # 找到所有可见的 ant-modal-content
    modals = browser.find_elements(By.XPATH, "//div[contains(@class,'ant-modal-content') and not(contains(@style,'display: none'))]")
    modal = None
    for m in modals:
        html = m.get_attribute("outerHTML")
        soup = BeautifulSoup(html, 'html.parser')
        h1 = soup.find('h1')
        if h1 and expected_title in h1.get_text(strip=True):
            modal = soup
            break
    if not modal:
        print("未找到详情弹窗，返回空")
        return {}
    info = {}
    h1 = modal.find('h1')
    info['标题'] = h1.get_text(strip=True) if h1 else ''
    body = modal.find('div', class_='ant-modal-body') if modal else None
    if body:
        code = body.find(string=lambda s: s and '条目代码' in s)
        info['条目代码'] = code.find_next().get_text(strip=True) if code else ''
        status = body.find(string=lambda s: s and '条目状态' in s)
        info['条目状态'] = status.find_next().get_text(strip=True) if status else ''
        permit = body.find(string=lambda s: s and '许可情况' in s)
        info['许可情况'] = permit.find_next().get_text(strip=True) if permit else ''
        desc = body.find(string=lambda s: s and '说明' in s)
        info['说明'] = desc.find_parent('div').get_text(strip=True) if desc else ''
        rel = body.find(string=lambda s: s and '相关活动' in s)
        info['相关活动'] = rel.find_parent('div').get_text(strip=True) if rel else ''
        norel = body.find(string=lambda s: s and '非相关活动' in s)
        info['非相关活动'] = norel.find_parent('div').get_text(strip=True) if norel else ''
        note = body.find(string=lambda s: s and '注' in s)
        info['备注'] = note.find_parent('div').get_text(strip=True) if note else ''
    else:
        info['条目代码'] = info['条目状态'] = info['许可情况'] = info['说明'] = info['相关活动'] = info['非相关活动'] = info['备注'] = ''
    return info

# 进入网站后，先关闭"关 闭"弹窗
def close_first_popup(browser):
    try:
        WebDriverWait(browser, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[span[contains(text(),'关') and contains(text(),'闭')]]"))
        )
        close_btns = browser.find_elements(By.XPATH, "//button[span[contains(text(),'关') and contains(text(),'闭')]]")
        for btn in close_btns:
            if btn.is_displayed() and btn.is_enabled():
                browser.execute_script("arguments[0].click();", btn)
                print("已自动点击关 闭按钮")
                break
    except Exception as e:
        print("首次弹窗未出现或已关闭，无需处理")

# 每次抓取详情后，关闭详情弹窗
def close_detail_popup(browser):
    # 滚动弹窗内容到底部
    try:
        modal_body = browser.find_element(By.CSS_SELECTOR, ".ant-modal-body")
        browser.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", modal_body)
        time.sleep(0.5)
    except Exception as e:
        print("滚动弹窗失败，忽略：", e)
    # 再次等待"取 消"按钮出现
    for _ in range(10):
        cancel_btns = browser.find_elements(By.XPATH, "//button[span[contains(text(),'取') and contains(text(),'消')]]")
        visible_btns = [btn for btn in cancel_btns if btn.is_displayed() and btn.is_enabled()]
        if visible_btns:
            browser.execute_script("arguments[0].click();", visible_btns[0])
            print("已自动点击取 消按钮")
            return
        time.sleep(1)
    print("警告：未找到可交互的取 消按钮！")

def jump_to_page(driver, page_num):
    # 滚动到底部
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(0.5)
    # 定位输入框
    input_box = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "input[aria-label='页']"))
    )
    # 先点击激活
    input_box.click()
    time.sleep(0.2)
    # 清空输入框（多次尝试，防止清空失败）
    input_box.clear()
    input_box.send_keys(Keys.CONTROL, 'a')
    input_box.send_keys(Keys.BACKSPACE)
    time.sleep(0.1)
    # 再输入页码
    input_box.send_keys(str(page_num))
    time.sleep(0.1)
    # 按下 Enter
    input_box.send_keys(Keys.ENTER)
    # 等待页面跳转
    WebDriverWait(driver, 10).until(
        EC.text_to_be_present_in_element((By.CSS_SELECTOR, "li.ant-pagination-item-active"), str(page_num))
    )

if __name__ == "__main__":
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        browser = None
        try:
            print(f"第 {retry_count + 1} 次尝试")
            browser = init_browser()
            
            browser.get(f"https://jyfwyun.com/search/item?keyword={SEARCH_KEYWORD}")
            print("已访问搜索结果页")
            
            # 等待页面完全加载
            time.sleep(2.5)  # 提高速度
            
            close_first_popup(browser)
            if not select_region(browser):
                print("地区选择失败，重试...")
                retry_count += 1
                continue
                
            time.sleep(1.5)  # 提高速度
            
            jump_to_page(browser, start_page)
            print(f"已跳转到第 {start_page} 页")

            all_data = []
            import random
            page_num = start_page
            max_pages = end_page
            while page_num <= max_pages:
                print(f"开始处理第 {page_num} 页")
                # 获取栏目列表
                columns = browser.find_elements(By.XPATH, "//span[contains(@class, 'zl-jyfw-item-text-color')]")
                titles = [col.text for col in columns]
                print(f"第{page_num}页找到 {len(titles)} 个栏目: {titles}")
                
                for i in range(min(5, len(titles))):  # 每页只处理5条
                    try:
                        title = titles[i]
                        unique_key = f"{title}_{page_num}"
                        if unique_key in existing:
                            print(f"栏目 {title} 第{page_num}页 已存在，跳过")
                            continue
                        print(f"正在处理栏目 {i+1}/5: {title}")
                        
                        # 重新获取栏目列表
                        columns = browser.find_elements(By.XPATH, "//span[contains(@class, 'zl-jyfw-item-text-color')]")
                        if i >= len(columns):
                            print(f"栏目索引超出范围，跳过")
                            continue
                        col = columns[i]
                        
                        # 滚动到元素位置
                        browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", col)
                        time.sleep(1)  # 提高速度
                        
                        # 模拟人类点击行为
                        ActionChains(browser).move_to_element(col).pause(0.5).click().perform()
                        print(f"已点击栏目: {title}")
                        
                        # 等待模态框
                        try:
                            WebDriverWait(browser, 7).until(
                                EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'ant-modal-content')]"))
                            )
                            time.sleep(1.5)  # 提高速度
                            
                            # 提取数据
                            modal_info = extract_modal_info_selenium(browser, title)
                            if modal_info and any(modal_info.values()):
                                modal_info['栏目名称'] = title
                                modal_info['页码'] = page_num
                                all_data.append(modal_info)
                                existing.add(unique_key)
                                print(f"已提取栏目 {title} 的数据: {modal_info}")
                            else:
                                print(f"栏目 {title} 数据提取失败或为空")
                            
                        except TimeoutException:
                            print(f"栏目 {title} 点击后未出现模态框")
                        
                        # 关闭模态框
                        try:
                            close_detail_popup(browser)
                            time.sleep(random.uniform(1.5, 2.5))  # 提高速度
                        except:
                            print("关闭模态框失败")
                        
                    except Exception as e:
                        print(f"处理栏目 {title} 时出错: {e}")
                        # 检查浏览器是否还在运行
                        try:
                            browser.current_url
                        except:
                            print("浏览器会话已断开，重新启动")
                            break
                        continue
                # 翻页逻辑
                try:
                    next_btn = browser.find_element(By.XPATH, "//li[@title='下一页' and not(contains(@class, 'ant-pagination-disabled'))]")
                    browser.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_btn)
                    time.sleep(0.5)  # 提高速度
                    next_btn.click()
                    print(f"已点击第 {page_num} 页的下一页")
                    page_num += 1
                    time.sleep(2)  # 提高速度
                except Exception as e:
                    print("没有更多页面，抓取结束")
                    break
            # 保存数据，合并去重
            if all_data:
                df_new = pd.DataFrame(all_data)
                if not df_exist.empty:
                    df_all = pd.concat([df_exist, df_new], ignore_index=True)
                    df_all.drop_duplicates(subset=['栏目名称', '页码'], inplace=True)
                    df_all.to_csv(csv_path, encoding="utf-8-sig", index=False)
                else:
                    df_new.to_csv(csv_path, encoding="utf-8-sig", index=False)
                print(f"已保存 {len(all_data)} 条新数据")
                break
            else:
                print("未提取到任何数据，重试...")
                retry_count += 1
                
        except Exception as e:
            print(f"程序运行出错: {e}")
            retry_count += 1
        finally:
            if browser:
                try:
                    browser.quit()
                except:
                    pass
            if retry_count < max_retries:
                print(f"等待 10 秒后重试...")
                time.sleep(10)
    
    if retry_count >= max_retries:
        print("已达到最大重试次数，程序退出")