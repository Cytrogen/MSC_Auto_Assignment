from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import time


def wait_for_element(driver, by, value, timeout=20):
    """等待元素出现并返回"""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def wait_for_elements(driver, by, value, timeout=20):
    """等待多个元素出现并返回"""
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_all_elements_located((by, value))
    )


def wait_for_manual_login():
    """等待用户手动登录"""
    input("请完成登录后按回车键继续...")


def auto_login(driver, email, password):
    """程序会自动登录"""
    try:
        print("开始自动登录...")

        # 等待email输入框出现并输入
        email_input = wait_for_element(driver, By.XPATH, '//*[@id="login-email"]')
        email_input.clear()
        email_input.send_keys(email)
        print("已输入邮箱")

        # 等待password输入框出现并输入
        password_input = wait_for_element(driver, By.XPATH, '//*[@id="login-password"]')
        password_input.clear()
        password_input.send_keys(password)
        print("已输入密码")

        # 等待登录按钮出现并点击
        login_button = wait_for_element(driver, By.XPATH, '//*[@id="login-submit-btn"]')
        login_button.click()
        print("已点击登录按钮")

        # 等待登录完成
        time.sleep(5)

        # 检查登录是否成功（可以根据实际情况调整检查方式）
        if "login" not in driver.current_url.lower():
            print("登录成功！")
            return True
        else:
            print("登录可能失败，请检查页面状态")
            return False

    except Exception as e:
        print(f"登录过程中发生错误: {str(e)}")
        return False


def wait_for_page_load(driver):
    """等待页面加载完成"""
    print("等待页面加载...")
    try:
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        time.sleep(5)
        return True
    except TimeoutException:
        print("页面加载超时")
        return False


def switch_to_iframe(driver, iframe_id):
    """切换到指定的iframe"""
    try:
        print(f"切换到iframe: {iframe_id}...")
        # 等待iframe出现
        iframe = wait_for_element(driver, By.ID, iframe_id)
        # 切换到iframe
        driver.switch_to.frame(iframe)
        return True
    except Exception as e:
        print(f"切换到iframe {iframe_id} 时发生错误: {str(e)}")
        return False


def switch_to_default_content(driver):
    """切换回主文档"""
    try:
        driver.switch_to.default_content()
        return True
    except Exception as e:
        print(f"切换回主文档时发生错误: {str(e)}")
        return False


def find_assignments(driver):
    """查找作业列表"""
    try:
        # 确保已经切换到正确的iframe
        if not switch_to_iframe(driver, "iframeContent"):
            return None

        print("查找所有作业...")
        assignments = wait_for_elements(
            driver,
            By.XPATH,
            "//li[contains(@data-section-id, '149791932')]"
        )

        print(f"找到 {len(assignments)} 个作业")
        return assignments

    except Exception as e:
        print(f"查找作业时发生错误: {str(e)}")
        return None
    finally:
        # 切换回主文档
        switch_to_default_content(driver)


def check_attempt_number(driver):
    """检查是否为第一次尝试"""
    try:
        # 切换到lti_iframe
        if not switch_to_iframe(driver, "lti_iframe"):
            return False

        # 查找attempt按钮
        attempt_button = wait_for_element(
            driver,
            By.CSS_SELECTOR,
            'a.cnButton.cnButton--primary.fr'
        )

        # 获取按钮文本并检查attempt次数
        button_text = attempt_button.text
        print("attempt次数：", button_text)
        if "Continue attempt 1" not in button_text:
            print("不是第一次尝试，退出当前作业...")
            # 查找并点击退出按钮
            exit_button = wait_for_element(
                driver,
                By.CSS_SELECTOR,
                'a.cnButton.cnButton--default'
            )
            exit_button.click()
            return False

        return True
    except Exception as e:
        print(f"检查attempt次数时发生错误: {str(e)}")
        return False
    finally:
        switch_to_default_content(driver)


def wait_for_submit_button(driver, timeout=500):
    """等待Submit按钮出现并点击"""
    print("等待Submit按钮出现...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # 寻找Submit按钮
            submit_button = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.XPATH, "//button[text()='Submit']"))
            )
            print("找到Submit按钮，准备点击")
            submit_button.click()
            print("已点击Submit按钮")
            return True
        except:
            # 如果没找到，切换回mzl-play-iframe继续等待
            try:
                driver.switch_to.frame("lti_iframe")
                driver.switch_to.frame("mzl-play-iframe")
            except:
                pass
            time.sleep(2)
            print("\r等待音频播放完成...", end='')

    print("\n等待Submit按钮超时")
    return False


def process_assignments(driver):
    try:
        # 等待页面完全加载
        if not wait_for_page_load(driver):
            return

        successful = 0
        failed = 0
        current_index = 0

        while True:
            # 每次循环重新获取作业列表
            assignments = find_assignments(driver)
            if not assignments:
                print("未找到任何作业")
                break

            if current_index >= len(assignments):
                break

            print(f"\n正在处理第 {current_index + 1}/{len(assignments)} 个作业")
            try:
                if process_assignment(driver, assignments[current_index]):
                    successful += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"处理作业时发生错误: {str(e)}")
                failed += 1

            current_index += 1
            # 每次处理完一个作业后等待页面稳定
            time.sleep(3)

        print(f"\n作业处理完成:")
        print(f"成功: {successful}")
        print(f"失败: {failed}")

    except Exception as e:
        print(f"处理作业列表时发生错误: {str(e)}")


def process_assignment(driver, assignment):
    """处理单个作业"""
    try:
        # 切换到iframe
        if not switch_to_iframe(driver, "iframeContent"):
            return False

        try:
            # 获取作业名称
            assignment_name = assignment.find_element(By.TAG_NAME, 'h3').text
            print(f"\n处理作业: {assignment_name}")

            # 滚动到作业元素并等待元素可见
            driver.execute_script("arguments[0].scrollIntoView(true);", assignment)
            time.sleep(2)

            # 点击打开侧边栏的按钮
            button = WebDriverWait(driver, 10).until(
                lambda x: assignment.find_element(
                    By.XPATH,
                    './/button[@class="assignment-actions text-right"]'
                )
            )
            driver.execute_script("arguments[0].click();", button)
            print("已点击打开侧边栏按钮")

        except Exception as e:
            print(f"处理作业初始步骤时发生错误: {str(e)}")
            return False

        # 等待侧边栏加载
        time.sleep(2)

        # 使用正确的ID选择器查找开始按钮
        print("正在查找启动按钮...")
        try:
            launch_button = wait_for_element(driver, By.ID, "launchRHP")
            print("找到启动按钮，准备点击")

            driver.execute_script("arguments[0].click();", launch_button)
            print("点击启动按钮完成")

            time.sleep(2)
        except Exception as e:
            print(f"查找或点击启动按钮时出错: {str(e)}")
            return False

        # 切换回主文档
        switch_to_default_content(driver)

        # 检查是否为第一次尝试
        if not check_attempt_number(driver):
            return False

        # 重新切换到lti_iframe
        if not switch_to_iframe(driver, "lti_iframe"):
            return False

        # 点击开始attempt按钮
        start_attempt = wait_for_element(
            driver,
            By.CSS_SELECTOR,
            'a.cnButton.cnButton--primary.fr'
        )
        start_attempt.click()
        print("已点击开始attempt按钮")
        time.sleep(2)

        # 注意：此时我们在lti_iframe中，不需要切换回主文档
        # 等待mzl-play-iframe在lti_iframe中加载
        print("等待mzl-play-iframe加载...")
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "mzl-play-iframe"))
            )
            time.sleep(2)
        except TimeoutException:
            print("等待mzl-play-iframe超时")
            return False

        # 从lti_iframe切换到mzl-play-iframe
        try:
            iframe = wait_for_element(driver, By.ID, "mzl-play-iframe")
            driver.switch_to.frame(iframe)
            print("已切换到mzl-play-iframe")
        except Exception as e:
            print(f"切换到mzl-play-iframe时发生错误: {str(e)}")
            return False

        # 等待并点击begin按钮
        try:
            begin_button = wait_for_element(
                driver,
                By.XPATH,
                "//button[text()='begin']"
            )
            begin_button.click()
            print("已点击begin按钮")
        except Exception as e:
            print(f"点击begin按钮时发生错误: {str(e)}")
            # return False

        # 等待Submit按钮出现并点击
        if not wait_for_submit_button(driver):
            print("等待Submit按钮失败，跳过当前作业")
            return False

        time.sleep(2)

        # 切换回主文档
        switch_to_default_content(driver)

        # 切换到lti_iframe
        if not switch_to_iframe(driver, "lti_iframe"):
            return False

        # 点击关闭按钮
        try:
            close_button = wait_for_element(
                driver,
                By.XPATH,
                '//*[@id="mzl-cn-header"]/div[2]/button'
            )
            close_button.click()
            print("已点击关闭按钮")
            time.sleep(3)
        except Exception as e:
            print(f"点击关闭按钮时发生错误: {str(e)}")
            return False

        print(f"完成作业: {assignment_name}")
        return True

    except Exception as e:
        print(f"\n处理作业时发生错误: {str(e)}")
        print("尝试关闭当前作业...")
        try:
            close_button = driver.find_element(
                By.CSS_SELECTOR,
                'button[aria-label="Close Activity"]'
            )
            close_button.click()
        except:
            pass
        time.sleep(3)
        return False
    finally:
        # 确保切换回主文档
        switch_to_default_content(driver)


def switch_to_nested_iframe(driver, parent_frame_id, child_frame_id):
    """切换到嵌套的iframe"""
    try:
        print(f"切换到父iframe: {parent_frame_id}...")
        # 首先切换到父iframe
        parent_frame = wait_for_element(driver, By.ID, parent_frame_id)
        driver.switch_to.frame(parent_frame)

        print(f"切换到子iframe: {child_frame_id}...")
        # 然后切换到子iframe
        child_frame = wait_for_element(driver, By.ID, child_frame_id)
        driver.switch_to.frame(child_frame)

        return True
    except Exception as e:
        print(f"切换到嵌套iframe时发生错误: {str(e)}")
        return False


def main():
    # 设置Edge选项
    options = webdriver.EdgeOptions()
    options.add_argument('--start-maximized')

    # 初始化WebDriver
    driver = webdriver.Edge(options=options)

    try:
        # 访问页面
        url = "https://newconnect.mheducation.com/"
        print("正在访问页面...")
        driver.get(url)

        # 自动登录或手动登录选择
        login_mode = input("选择登录方式 (1: 自动登录, 2: 手动登录): ")

        if login_mode == "1":
            # 自动登录，自行填写.env
            load_dotenv()

            email = os.getenv("EMAIL")
            password = os.getenv("PASSWORD")
            if not auto_login(driver, email, password):
                print("自动登录失败，程序退出")
                return
        else:
            # 手动登录
            print("请先登录您的账号...")
            input("完成登录后按回车键继续...")

        # 给予额外时间确保登录后页面完全加载
        time.sleep(5)

        # 处理所有作业
        process_assignments(driver)

    except Exception as e:
        print(f"发生错误: {str(e)}")

    finally:
        # 询问是否关闭浏览器
        if input("是否关闭浏览器？(y/n): ").lower() == 'y':
            driver.quit()
            print("浏览器已关闭")
        else:
            print("请手动关闭浏览器窗口")


if __name__ == "__main__":
    main()
