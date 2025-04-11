from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
import time
import random
import ddddocr  # 导入 ddddocr 库用于验证码识别

# 初始化 WebDriver 和 ddddocr
edge_service = Service("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe")
driver = webdriver.Edge(service=edge_service)
ocr = ddddocr.DdddOcr()  # 创建 OCR 对象

def select_dropdown_by_name(driver, dropdown_name, option_value):
    """
    通用函数：选择指定名称的下拉框并设置值
    :param driver: WebDriver 实例
    :param dropdown_name: 下拉框的 name 属性
    :param option_value: 要选择的选项值
    """
    try:
        # 定位下拉框
        dropdown = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, dropdown_name))
        )
        print(f"找到下拉框: {dropdown_name}")

        # 设置 <select> 的值
        driver.execute_script(f"document.getElementsByName('{dropdown_name}')[0].value = '{option_value}';")
        selected_value = driver.execute_script(f"return document.getElementsByName('{dropdown_name}')[0].value;")
        print(f"已选择下拉框 {dropdown_name}: 值为 {selected_value}")
    except Exception as e:
        print(f"选择下拉框 {dropdown_name} 时出错: {e}")
        raise

try:
    # ===================第一步===================
    driver.get("https://bsdt.zjtongji.edu.cn/taskcenter/workflow/todo")  # 目标系统的登录 URL
    time.sleep(2)  # 等待页面加载

    # 输入用户名
    username_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    username_input.send_keys("z20220230804")  # 替换为实际用户名

    # 输入密码
    password_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    password_input.send_keys("gprbEdYf7xt3!V5")  # 替换为实际密码

    # 自动识别验证码
    captcha_image_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "veriyCodeImg"))  # 定位验证码图片元素
    )
    captcha_image_element.screenshot("captcha.png")  # 截图保存验证码图片
    print("验证码图片已保存为 captcha.png")

    with open("captcha.png", "rb") as f:
        img_bytes = f.read()
    captcha_text = ocr.classification(img_bytes)  # 使用 ddddocr 进行验证码识别
    print(f"识别到的验证码: {captcha_text}")

    # 输入验证码
    captcha_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "veriyCode"))
    )
    captcha_input.send_keys(captcha_text.strip())  # 输入识别结果

    # 点击登录按钮
    login_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[@onclick='return login()']"))
    )
    login_button.click()
    print("已尝试登录")
    time.sleep(5)  # 等待登录完成

    # ===================第二步===================
    while True:
        try:
            # 查找页面中的第一个“办理”按钮
            handle_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='sx_item_but padding_left']//input[@value='办理']"))
            )
            print("找到办理按钮")
            handle_button.click()
            print("已点击办理按钮，新标签页将打开审批页面")
            time.sleep(5)  # 等待新标签页加载

            # 获取所有窗口句柄
            window_handles = driver.window_handles
            if len(window_handles) > 1:
                # 切换到新标签页
                driver.switch_to.window(window_handles[-1])
                print("已切换到新标签页")
                time.sleep(5)  # 等待新标签页加载完成
            else:
                print("未检测到新标签页，跳过本次循环")
                continue

            if random.random() < 0.75:  # 75% 的概率通过
                print("选择同意操作")
                operation = "同意"
            else:  # 25% 的概率退回
                print("选择退回操作")
                operation = "退回"

            # 1. 处理“是否与家长联系”下拉框
            contact_parent_option_value = random.choice(["1", "2"])  # 1 对应"是", 2 对应"否"
            select_dropdown_by_name(driver, "fieldBZRSFLX", contact_parent_option_value)

            # 2. 处理“常用意见”下拉框
            if operation == "同意":
                common_opinion_option_value = "1"
                suggestion_text = "同意"
            else:
                common_opinion_option_value = "4"
                suggestion_text = "不同意"
            select_dropdown_by_name(driver, "fieldBZRGeneral", common_opinion_option_value)

            # 3. 输入文字建议
            suggestion_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "fieldBZROpinion"))  # 使用 name 定位文本框
            )
            suggestion_input.clear()
            suggestion_input.send_keys(suggestion_text)
            print(f"已输入文字建议: {suggestion_text}")
            time.sleep(3)

            # 4. 根据操作类型点击对应的按钮
            if operation == "同意":
                # 查找并点击“同意”按钮
                driver.execute_script("""
                    const button = document.querySelector('a[name="infoplus_action_5557_155"]');
                    if (button) {
                        button.click();
                        console.log("已点击同意按钮");
                    } else {
                        console.error("未找到同意按钮");
                    }
                """)
                print("已通过 JavaScript 点击同意按钮")
            else:
                # 查找并点击“退回”按钮
                driver.execute_script("""
                    const button = document.querySelector('a[name="infoplus_action_5557_93"]');
                    if (button) {
                        button.click();
                        console.log("已点击退回按钮");
                    } else {
                        console.error("未找到退回按钮");
                    }
                """)
                print("已通过 JavaScript 点击退回按钮")

            print(f"已{operation}该假单")

            time.sleep(5)  # 等待提交完成

            # 关闭当前标签页并返回主标签页
            driver.close()  # 关闭当前标签页
            driver.switch_to.window(window_handles[0])  # 切换回主标签页
            print("已关闭审批页面并返回假单列表")
            time.sleep(3)  # 等待页面重新加载

        except Exception as e:
            print(f"处理假单时出错: {e}")
            import traceback
            traceback.print_exc()  # 打印完整的堆栈信息
            break  # 如果出现错误，退出循环

except Exception as e:
    print("发生错误:", e)

finally:
    # 关闭浏览器
    driver.quit()