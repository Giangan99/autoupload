from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
import time
import openpyxl

chrome_path = r'F:\GoogleChromePortable\chromedriver.exe'
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = r'F:\GoogleChromePortable\App\Chrome-bin\chrome.exe'

# Thay đổi đường dẫn đến thư mục profile của Chrome Portable với tên Profile nhập vào
chrome_options.add_argument("--profile-directory=Profile 10")

chrome_options.add_argument("--user-data-dir=F:\\GoogleChromePortable\\Data\\profile\\")

service = Service(chrome_path)

driver = webdriver.Chrome(service=service, options=chrome_options)

# Đường dẫn đến thư mục chứa tệp tin ảnh
folder_path = input("Nhập đường dẫn file ảnh: ")

def get_file_paths(folder_path):
    image_paths = []
    excel_paths = []

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
            image_paths.append(file_path)
        elif filename.lower().endswith(('.xlsx', '.xls')):
            excel_paths.append(file_path)

    return image_paths, excel_paths
image_paths, excel_paths = get_file_paths(folder_path)
print("Image paths:")
for image_path in image_paths:
    print(image_path)

print("Excel paths:")
for excel_path in excel_paths:
    print(excel_path)

for i in range(len(image_paths)):
    wait = WebDriverWait(driver, 10)
    driver.get("https://www.redbubble.com/portfolio/images/new?ref=dashboard")
    # Click vào upload_button
    button = driver.find_element_by_id(select-image-single)
    button.click()

    # Tìm phần tử input file
    input_file = driver.find_element(By.CSS_SELECTOR, "input[type='file']")

    # Gửi đường dẫn ảnh
    input_file.send_keys(image_paths[i])

    # Chờ một khoảng thời gian để tải lên ảnh
    time.sleep(5)
    # Click vào lựa chọn "No"
    no_option = driver.find_element(By.ID,"design_content_flag_false")
    no_option.click()

    # Gửi đường dẫn file Excel
    excel_file_path = excel_paths[i]
    # Load file Excel
    workbook = openpyxl.load_workbook(excel_file_path)
    # Chọn sheet trong file Excel
    worksheet = workbook['Sheet1']

    # Đọc dữ liệu từ file Excel
    title = worksheet['A2'].value
    main_tags = worksheet['B2'].value
    describe = worksheet['C2'].value

    # Đóng file Excel
    workbook.close()

    title_input = driver.find_element(By.ID, "design_design_title")
    title_input.clear()  # Xóa nội dung hiện tại trong trường dữ liệu
    title_input.send_keys(title)  # Điền nội dung mới vào trường dữ liệu

    time.sleep(2)
    main_tags_input = driver.find_element(By.CLASS_NAME, 'form__control.ui-autocomplete-input')
    main_tags_input.send_keys(main_tags)
    time.sleep(2)
    describe_input = driver.find_element(By.CLASS_NAME, 'form__control.textarea')
    describe_input.send_keys(describe)
    time.sleep(2)
    # Tìm element bằng class "taggle_placeholder"
    placeholder_span = driver.find_element(By.CLASS_NAME, "taggle_placeholder")
    placeholder_span.click()  # Click vào thẻ <span>
    time.sleep(2)

    # Lặp qua từng dòng trong file Excel
    for row in worksheet.iter_rows(min_row=2, max_row=15, values_only=True):
        supporting_tags = row[3]

        if supporting_tags is not None:
            supporting_tags_input = driver.find_element(By.CLASS_NAME, "taggle_input.ui-autocomplete-input")
            if supporting_tags_input is not None:
                supporting_tags_input.clear()
                supporting_tags_input.send_keys(supporting_tags[:14])
                supporting_tags_input.send_keys(Keys.RETURN)         

    driver.switch_to.active_element.send_keys(Keys.TAB)
    time.sleep(5)
    # Click vào lựa chọn màu
    color_dropdown = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "dd-selected-text")))
    color_dropdown.click()
    # Tìm lựa chọn màu trắng và nhấp vào nó
    white_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//label[@class='dd-option-text' and text()='White']")))
    white_option.click()
    # Tìm tất cả các thẻ <div> có lớp "canvas_label"
    canvas_labels = WebDriverWait(driver, 20).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "div.canvas_label")))

    # Lặp qua danh sách các thẻ và click vào thẻ mong muốn
    for canvas_label in canvas_labels:
        if canvas_label.text.strip() == "Hoodie":
            canvas_label.click()
            break  # Kết thúc vòng lặp sau khi click thành công
    # Tìm phần tử bằng XPath
    color_dropdown1 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="primary_color_hoodie"]/div/a/label')))

    # Click vào phần tử
    color_dropdown1.click()

    # Tìm lựa chọn màu theo tên (ví dụ: "Vintage Heather") và nhấp vào nó
    color_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//label[@class='dd-option-text' and text()='Vintage Heather']")))
    color_option.click()
    # Tìm tất cả các thẻ <div> có lớp "canvas_label"
    canvas_labels = WebDriverWait(driver, 10).until(EC.visibility_of_all_elements_located((By.CSS_SELECTOR, "div.canvas_label")))

    # Lặp qua danh sách các thẻ và click vào thẻ mong muốn
    for canvas_label in canvas_labels:
        if canvas_label.text.strip() == "Baseball Tee":
            canvas_label.click()
            break  # Kết thúc vòng lặp sau khi click thành công
    # Tìm phần tử bằng XPath
    color_dropdown2 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="primary_color_baseballtee"]/div/a/label')))

    # Click vào phần tử
    color_dropdown2.click()

    # Tìm lựa chọn màu theo tên (ví dụ: "Vintage Heather") và nhấp vào nó
    color_option1 = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//label[@class='dd-option-text' and text()='White/Black']")))
    color_option1.click()

    chexbox = driver.find_element(By.ID,"terms")
    chexbox.click()

    # Click vào đăng
    publish_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.publish-and-promote-button.btn.btn--big.btn--green")))

    # Thực hiện click vào nút
    publish_button.click()
driver.quit()


