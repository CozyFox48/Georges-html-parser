from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pyautogui
import excel_handler as EH
import time
import os

def force(callback, debug=False):
	while True:
		time.sleep(0.5)
		try:
			callback()
			break
		except:
			pass

chrome_options = Options()
chrome_options.add_argument("--save-page-as-mhtml")
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=chrome_options)
driver.get('http://www.wm.tu-darmstadt.de/mat-db/signin.php')

login_input = driver.find_element(By.NAME, 'login')
password_input = driver.find_element(By.NAME, 'password')

login_input.clear()
login_input.send_keys('Guest')
password_input.clear()
password_input.send_keys('FG_WM')

signin_button = driver.find_elements(By.TAG_NAME, 'input')[2]
signin_button.click()

driver.get('http://www.wm.tu-darmstadt.de/mat-db/search.php')

name_select = driver.find_element(By.NAME, 'name2')

html_content = driver.execute_script('return arguments[0].innerHTML', name_select)
soup = BeautifulSoup(html_content, 'html.parser')
name_list = soup.get_text().split('\n')[2:-1]

row_number = 4
id_list = []

for name2 in name_list:

	name_select = driver.find_element(By.NAME, 'name2')

	force(lambda:name_select.click())
	force(lambda:driver.find_elements(By.XPATH, f"//option[text()='{name2}']")[0].click())
	force(lambda:name_select.click())
	force(lambda:driver.find_elements(By.XPATH, f"//input[@value='search']")[0].click())

	view_buttons = driver.find_elements(By.XPATH, f"//a/b[text()='view']")

	for view_button in view_buttons:

		view_url = driver.execute_script('return arguments[0].parentNode.href', view_button)
		_ID = view_url.split('id=')[-1]

		if _ID not in id_list:

			row_number = row_number + 1
			
			id_list.append(_ID)

			driver.execute_script("window.open('about:blank', '_blank');")
			driver.switch_to.window(driver.window_handles[-1])
			driver.get(view_url)

			content_stage = driver.find_elements(By.XPATH, f"//center/table/tbody")[0]

			first_content = driver.execute_script('return arguments[0].children[1].firstElementChild.firstElementChild.firstElementChild.firstElementChild', content_stage)
			last_content = driver.execute_script('return arguments[0].children[5].firstElementChild.firstElementChild.lastElementChild', content_stage)

			name_line = driver.execute_script('return arguments[0].children[1].innerHTML', first_content)
			soup = BeautifulSoup(name_line, 'html.parser')
			NAME = soup.get_text().replace('\xa0', '').split('\n')[3:-3][0]
			if NAME == '':
				NAME = soup.get_text().replace('\xa0', '').split('\n')[3:-3][2]
			NUMBER_TEXTUAL = soup.get_text().replace('\xa0', '').split('\n')[3:-3][1]

			LINK_VAL = NAME + ' - ' + _ID

			row_values = [LINK_VAL, int(_ID), NAME, NUMBER_TEXTUAL]

			driver.execute_cdp_cmd("Page.enable", {})
			result = driver.execute_cdp_cmd("Page.captureSnapshot", {"format": "mhtml"})

			with open('reference/' + LINK_VAL + '.mhtml', "wb") as file:
			    file.write(result["data"].encode('utf-8'))

			for _ in range(5):
				row_values.append(driver.execute_script(f'return arguments[0].children[{_+3}].lastElementChild.textContent', first_content))

			chemical_content = driver.execute_script(f'return arguments[0].children[9].firstElementChild.lastElementChild.firstElementChild.innerHTML', first_content)
			soup = BeautifulSoup(chemical_content, 'html.parser')
			chemical_temp = list(filter(lambda item: item != '', soup.get_text().replace('\xa0', '').split('\n')))
			if len(chemical_temp) == 1:
				chemical_temp.append('')
			chemical_composition = list(zip(chemical_temp[:int(len(chemical_temp)/2)], chemical_temp[int(len(chemical_temp)/2):]))

			microstructure_content = driver.execute_script('return arguments[0].children[11].firstElementChild.firstElementChild.firstElementChild', first_content)
			hardness_content = driver.execute_script('return arguments[0].children[11].lastElementChild.firstElementChild.firstElementChild', first_content)

			row_values.append(driver.execute_script('return arguments[0].children[1].firstElementChild.innerHTML.replace("&nbsp;", "")', microstructure_content))
			row_values.append(driver.execute_script('return arguments[0].children[1].firstElementChild.innerHTML.replace("&nbsp;", "")', hardness_content))

			for _ in range(6):
				row_values.append(driver.execute_script(f'return arguments[0].children[{_+4}].lastElementChild.textContent', microstructure_content))

			for _ in range(10):
				row_values.append(driver.execute_script(f'return arguments[0].children[{_+4}].children[2].innerHTML.replace("&nbsp;", "")', hardness_content))

			for _ in range(13):
				row_values.append(driver.execute_script(f'return arguments[0].children[{_+16}].children[2].innerHTML.replace("&nbsp;", "")', hardness_content))

			for _ in range(11):
				row_values.append(driver.execute_script(f'return arguments[0].children[{_+16}].lastElementChild.innerHTML.replace("&nbsp;", "")', microstructure_content))

			row_values.append(driver.execute_script('return arguments[0].lastElementChild.lastElementChild.firstElementChild.firstElementChild.children[1].children[2].innerHTML.replace("&nbsp;", "")', first_content))
			row_values.append(driver.execute_script('return arguments[0].lastElementChild.lastElementChild.firstElementChild.firstElementChild.children[2].children[2].innerHTML.replace("&nbsp;", "")', first_content))
			
			row_values.append(driver.execute_script('return arguments[0].lastElementChild.firstElementChild.firstElementChild.firstElementChild.lastElementChild.textContent.trim()', first_content))

			experimental_parent = driver.execute_script('return arguments[0].innerHTML', last_content)
			soup = BeautifulSoup(experimental_parent, 'html.parser')
			experimental_temp = soup.get_text().replace('\xa0', '').split()
			experimental_data = [experimental_temp[_:_+5] for _ in range(0, len(experimental_temp), 5)]

			EH.append_summary(row_number, row_values)
			EH.add_sheet(LINK_VAL, chemical_composition, experimental_data)
			
			driver.close()
			driver.switch_to.window(driver.window_handles[0])

EH.remove_tabular()

driver.quit()