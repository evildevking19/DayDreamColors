import os, time, requests, sys
from io import BytesIO
import schedule
from PIL import Image
from colorama import Fore, init
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openai import OpenAI
from timeout import timeout
from constants import *

try:
    from ctypes import windll  # Only exists on Windows.
    myappid = 'Day Dream Colors Automation'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

gpt_client = OpenAI(api_key=OPENAI_KEY)
running_progress_num = -1

@timeout(30)
def getChatGPTResponse(prompt):
    response = gpt_client.chat.completions.create(
        model="gpt-3.5-turbo-16k",
        messages=[
            {"role": "user", "content": prompt}
        ]
    )
    return response

def addNewPost(data, images):
    print(f"{Fore.LIGHTCYAN_EX}Starting add new post on website...{Fore.RESET}")
    browser = getGoogleDriver()
    browser.get(NEW_POST_URL)
    
    print(f"{Fore.LIGHTCYAN_EX}Logging in...{Fore.RESET}")
    # Login part
    username_element = browser.find_element(By.XPATH, "//input[@id='user_login']")
    username_element.click()
    username_element.send_keys(USERNAME)
    
    password_element = browser.find_element(By.XPATH, "//input[@id='user_pass']")
    password_element.click()
    password_element.send_keys(PASSWORD)
    
    submit_element = browser.find_element(By.XPATH, "//input[@id='wp-submit']")
    submit_element.click()
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Logged in")
    
    # Add page title part
    title_element = browser.find_element(By.XPATH, "//div[@class='edit-post-visual-editor__post-title-wrapper']").find_element(By.TAG_NAME, "h1")
    title_element.click()
    title_element.send_keys(data["page_title"])
    title_element.send_keys(Keys.ENTER)
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Adding a new page title")
    
    # Add description of the page using ChatGPT output
    while True:
        try:
            response = getChatGPTResponse(f"{CHATGPT_PROMPT} : {data['page_title']}")
            new_text_block_element = browser.find_element(By.XPATH, "//p[@data-empty='true']")
            new_text_block_element.click()
            new_text_block_element.send_keys(response.choices[0].message.content)
            new_text_block_element.send_keys(Keys.ENTER)
            break
        except:
            print(f"{Fore.RED}Failed to response from ChatGPT. Retrying in 10 seconds...{Fore.RESET}")
            time.sleep(10)
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Adding page description from ChatGPT")
    time.sleep(1)
    
    # Add heading of the opening paragraph
    new_text_block_element = browser.find_element(By.XPATH, "//p[@data-empty='true']")
    new_text_block_element.click()
    time.sleep(1)
    add_block_btn = browser.find_element(By.XPATH, "//button[@aria-label='Add block']")
    add_block_btn.click()
    time.sleep(1)
    add_heading_btn = browser.find_element(By.XPATH, "//button[@class='components-button block-editor-block-types-list__item editor-block-list-item-heading']")
    add_heading_btn.click()
    time.sleep(1)
    heading_editbox = browser.find_element(By.XPATH, "//h2[@data-title='Heading']")
    heading_editbox.click()
    heading_editbox.send_keys(f"{data['page_title']} {data['seo_title']}")
    heading_editbox.send_keys(Keys.ENTER)
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Adding header of the opening paragraph")
    time.sleep(1)
    
    # Add each image with title
    for i, image_path in enumerate(images):
        print(f"{Fore.LIGHTCYAN_EX}Uploading:{Fore.RESET}    {image_path}...")
        new_text_block_element = browser.find_element(By.XPATH, "//p[@data-empty='true']")
        new_text_block_element.click()
        new_text_block_element.send_keys(f"image {i+1}")
        new_text_block_element.send_keys(Keys.ENTER)
        time.sleep(1)
        new_text_block_element = browser.find_element(By.XPATH, "//p[@data-empty='true']")
        ActionChains(browser).move_to_element_with_offset(new_text_block_element, 8, 0).perform()
        new_text_block_element.click()
        time.sleep(1.5)
        add_block_btn = browser.find_element(By.XPATH, "//button[@aria-label='Add block']")
        add_block_btn.click()
        time.sleep(1)
        add_image_btn = browser.find_element(By.XPATH, "//button[@class='components-button block-editor-block-types-list__item editor-block-list-item-image']")
        add_image_btn.click()
        time.sleep(1)
        upload_file = browser.find_element(By.XPATH, "//input[@type='file']")
        upload_file.send_keys(image_path)
        WebDriverWait(browser, 100).until(lambda browser: browser.execute_script('return document.querySelector("svg[class=\'components-spinner css-1yxc2ud ea4tfvq2\']")') == None)
        last_image_element = browser.find_element(By.XPATH, "//div[@class='components-resizable-box__container has-show-handle']")
        ActionChains(browser).move_to_element(last_image_element).click(last_image_element).send_keys(Keys.ENTER).perform()
        time.sleep(1)
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Image uploading")
    
    # Add SEO title
    seo_content = browser.find_elements(By.XPATH, "//div[@class='public-DraftStyleDefault-block public-DraftStyleDefault-ltr']")
    seo_title_element = seo_content[0]
    meta_desc_element = seo_content[1]
    
    ActionChains(browser).move_to_element(seo_title_element).click(seo_title_element).perform()
    seo_title_element.send_keys(Keys.CONTROL + "a")
    seo_title_element.send_keys(Keys.DELETE)
    seo_title_element.send_keys(data["page_title"] + data["seo_title"])
    
    ActionChains(browser).move_to_element(meta_desc_element).click(meta_desc_element).send_keys(data["meta_desc"]).perform()
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Adding SEO title and META description")
    
    # Add tags
    post_tab_element = browser.find_element(By.XPATH, "//button[@data-label='Post']")
    post_tab_element.click()
    tag_element = browser.find_element(By.XPATH, "//input[@class='components-form-token-field__input']")
    ActionChains(browser).move_to_element(tag_element).click(tag_element).perform()
    tags = data["tags"].split(",")
    for tag in tags:
        ActionChains(browser).click(tag_element).send_keys(tag.strip()).perform()
        time.sleep(1)
        tag_element.send_keys(Keys.ENTER)
        
    print(f"{Fore.GREEN}Successed:{Fore.RESET}    Adding tags")
    time.sleep(2)
    
    # Publish page
    publish_btn = browser.find_element(By.XPATH, "//button[@class='components-button editor-post-publish-panel__toggle editor-post-publish-button__button is-primary']")
    publish_btn.click()
    time.sleep(2)
    publish_btn_last = browser.find_element(By.XPATH, "//button[@class='components-button editor-post-publish-button editor-post-publish-button__button is-primary']")
    publish_btn_last.click()
    print(f"{Fore.LIGHTCYAN_EX}Publishing...{Fore.RESET}")
    time.sleep(3)
    
    browser.quit()

def downloadAIImage(prompt):
    images = []
    for i in range(3):
        try:
            response = gpt_client.images.generate(
                model="dall-e-3",
                prompt=prompt,
                size="1024x1024",
                quality="hd",
                n=1,
            )
            
            file_content = requests.get(response.data[0].url).content
            image = Image.open(BytesIO(file_content))
            image = image.resize((512,512),Image.Resampling.LANCZOS)
            filename = f"image{i+1}.jpg"
            image.save(filename, optimize=True, quality=95)
            
            images.append(os.path.abspath(filename))
            print(f"{Fore.GREEN}Downloaded:{Fore.RESET}    {filename}")
        except:
            break
    return images

def getGSData(sheetId):
    service = getGoogleService("sheets", "v4")
    result = service.spreadsheets().values().get(spreadsheetId=sheetId, range="Sheet1!A2:E").execute()
    result = result["values"]
    sheet_data = []
    for v in result:
        if len(v) == 0 or v[0].strip() == "": break
        data = {}
        data["prompt"] = v[0]
        data["page_title"] = v[1]
        data["seo_title"] = v[2]
        data["meta_desc"] = v[3]
        data["tags"] = v[4]
        sheet_data.append(data)
    return sheet_data

def runSchedule(sheet_data):
    global running_progress_num
    running_progress_num += 1
    if len(sheet_data) - 1 < running_progress_num:
        sys.exit()
    else:
        print(f"{Fore.LIGHTCYAN_EX}Downloading AI image...{Fore.RESET}")
        images = downloadAIImage(sheet_data[running_progress_num]["prompt"])
        if len(images) == 3:
            addNewPost(sheet_data[running_progress_num], images)
            print(f"{Fore.GREEN}Published a new post successfully!{Fore.RESET}")
        else:
            running_progress_num -= 1
            print(f"{Fore.RED}Failed to generate AI images. Retrying in 10 minutes...{Fore.RESET}")

if __name__ == "__main__":
    init()
    sheetId = input(f"{Fore.CYAN}Enter your google sheet id:{Fore.RESET}  ")
    if sheetId.strip() == "":
        print(f"{Fore.RED}Sorry, but your sheet id is not valid.{Fore.RESET}")
    else:
        # min = input(f"{Fore.CYAN}Enter a specific minutes for schedule:{Fore.RESET}  ")
        print(f"{Fore.LIGHTCYAN_EX}Fetching Google Sheet data...{Fore.RESET}")
        sheet_data = getGSData(sheetId)
        if len(sheet_data) == 0:
            print(f"{Fore.RED}Sorry, but there is no any sheet data.{Fore.RESET}")
        else:
            runSchedule(sheet_data)
            
            if len(sheet_data) != 1:
                # Run automation in every 10 min
                # schedule.every(int(min)).minutes.do(runSchedule, sheet_data)
                # Run automation in every 2 hours
                schedule.every(2).hours.do(runSchedule, sheet_data)
                
                while True:
                    schedule.run_pending()
                    time.sleep(1)
    print("Press any key to exit...")