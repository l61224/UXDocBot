import os, time, datetime
from TM1py import TM1Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import configparser

# ========================================================================
# Region - Definition
## Read information from config.ini (in currenct path)
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

## Screen_Catcher
UX_BASE_URL     = config['Screen_Catcher']['ux_base_url']       # UX Homepage URL/ UX首頁網址
SCREENSHOT_DIR  = config['Screen_Catcher']['screenshot_path']   # Screenshot storage location/ 截圖儲存位置
UX_APP_MDX      = config['Screen_Catcher']['ux_app_mdx']        # UX Apps You Want to Take Screenshots of/ 想要截圖的UX Apps (define by MDX)
DELAY           = int(config['Screen_Catcher']['page_delay'])   # Each app waits a few seconds before taking a screenshot/ 每個 App 等待幾秒後截圖

## UX_CS
USERNAME        = config['UX_CS']['login_username']             # Screenshot user login account/ 截圖使用者登入帳號
PASSWORD        = config['UX_CS']['login_password']             # Screenshot user login password/ 截圖使用者登入密碼
ADDRESS         = config['UX_CS']['address']                    # Content Store Address
PORT            = config['UX_CS']['port']                       # Content Store Port
NAMESPACE       = config['UX_CS']['namespace']                  # Content Store Namespace
SYS_USERNAME    = config['UX_CS']['system_username']            # Content Store Login System Account
SYS_PASSWORD    = config['UX_CS']['system_password']            # Content Store Login System Account Password
SSL             = config['UX_CS']['ssl']                        # Content Store Using SSL?

# Page auto-input location
USERNAME_INPUT_ID   = "tm1-login-user"
PASSWORD_INPUT_ID   = "tm1-login-pass"
LOGIN_BUTTON_ID     = "tm1-login-button"

# EndRegion - Definition
# ========================================================================

# ========================================================================
# Region - Get UX App List
## === TM1 Connection/ TM1 連線 ===
tm1 = TM1Service(address=ADDRESS, port=PORT, user=SYS_USERNAME, password=SYS_PASSWORD, ssl=SSL, namespace=NAMESPACE)
ux_apps = tm1.elements.execute_set_mdx_element_names( UX_APP_MDX)

# EndRegion - Get UX App List
# ========================================================================

# ========================================================================
# Region - Logging
# === Chrome Driver ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)

# === Log into UX/ 登入 UX ===
print("🔐 Try entering UX...")
driver.get(f"{UX_BASE_URL}/#/")

# === Wait for automatic login to complete/ 等待自動登入完成 === 
time.sleep(DELAY)

try:
    # === Wait for the screen to load the account input box (login screen appears)/ 等待畫面載入帳號輸入框（登入畫面出現） === 
    WebDriverWait(driver, DELAY).until(
        EC.presence_of_element_located((By.ID, USERNAME_INPUT_ID))
    )
    print("🔐 Enter account and password...")
    driver.find_element(By.ID, USERNAME_INPUT_ID).send_keys(USERNAME)
    driver.find_element(By.ID, PASSWORD_INPUT_ID).send_keys(PASSWORD)
    driver.find_element(By.ID, LOGIN_BUTTON_ID).click()

except TimeoutException:
    # === If the login screen does not appear, it may be because you are already logged in/ 如果登入畫面沒出現，可能是因為已登入 === 
    print("🔑 The login screen did not appear. You may have logged in automatically.")

# ===  Wait for the main screen element to confirm successful login/ 等待主畫面元素確認登入成功 === 
WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "main-wrapper"))
)
print("✅ Login successful, enter the main screen")

# EndRegion - Logging
# ========================================================================

# ========================================================================
# Region - Create Folder/ 建立資料夾
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# EndRegion - Create Folder/ 建立資料夾
# ========================================================================

# ========================================================================
# Region - Jump to the specified app & capture the screen/ 跳轉到指定 App & 截取畫面
for app_name in ux_apps:
    print(f"📸 ux_app: {app_name} screenshot completed！")
    ux_app_url_param = "#!/app/" + app_name
    
    # === Just change the hash and re-render the screen/ 只改 hash，讓畫面重新渲染 === 
    driver.execute_script(f"window.location.hash = '{ux_app_url_param}'")
    time.sleep(DELAY)  # Wait for page query/ 等頁面query

    safe_name = app_name.replace(" ", "_")
    driver.save_screenshot(f"{SCREENSHOT_DIR}/{safe_name}.png")

# EndRegion - Jump to the specified app & capture the screen/ 跳轉到指定 App & 截取畫面
# ========================================================================

print("✅ All App screenshots completed！")
driver.quit()
# === Sleep 1.5 sec for check result/ 等待1.5秒, 查看結果 ===
time.sleep(1.5)

