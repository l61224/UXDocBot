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
UX_BASE_URL     = config['Screen_Catcher']['ux_base_url']       # UX Homepage URL
SCREENSHOT_DIR  = config['Screen_Catcher']['screenshot_path']   # 截圖儲存位置
UX_APP_MDX      = config['Screen_Catcher']['ux_app_mdx']        # 想要截圖的UX Apps (define by MDX)
DELAY           = int(config['Screen_Catcher']['page_delay'])   # 每個 App 等待幾秒後截圖

## UX_CS
USERNAME        = config['UX_CS']['login_username']             # 截圖使用者登入帳號
PASSWORD        = config['UX_CS']['login_password']             # 截圖使用者登入密碼
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
## === TM1 連線 ===
# with TM1Service(address=ADDRESS, port=PORT, user=SYS_USERNAME, password=SYS_PASSWORD, ssl=SSL, namespace=NAMESPACE) as tm1:
 #   === 用 MDX 撈出 UX App 維度元素 === 
    # mdx     = UX_APP_MDX
    # ux_apps = tm1.dimensions.execute_mdx("}APQ UX App", mdx)
 
tm1 = TM1Service(address=ADDRESS, port=PORT, user=SYS_USERNAME, password=SYS_PASSWORD, ssl=SSL, namespace=NAMESPACE)
ux_apps = tm1.elements.execute_set_mdx_element_names( UX_APP_MDX)

# EndRegion - Get UX App List
# ========================================================================

# ========================================================================
# Region - Logging
# === 啟動 Chrome Driver ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chrome_options)

# === 登入 UX (SecurityMode = 1，可能會自動登入，不用填帳密) ===
print("🔐 嘗試進入 UX...")
driver.get(f"{UX_BASE_URL}/#/")

# === 等待自動登入完成 === 
time.sleep(DELAY)

try:
    # === 等待畫面載入帳號輸入框（登入畫面出現） === 
    WebDriverWait(driver, DELAY).until(
        EC.presence_of_element_located((By.ID, USERNAME_INPUT_ID))
    )
    print("🔐 自動輸入帳號密碼...")
    driver.find_element(By.ID, USERNAME_INPUT_ID).send_keys(USERNAME)
    driver.find_element(By.ID, PASSWORD_INPUT_ID).send_keys(PASSWORD)
    driver.find_element(By.ID, LOGIN_BUTTON_ID).click()

except TimeoutException:
    # === 如果登入畫面沒出現，可能是因為已登入 === 
    print("🔑 未出現登入畫面，可能已自動登入（SSO 或 session）")

# ===  等待主畫面元素確認登入成功 === 
WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.ID, "main-wrapper"))
)
print("✅ 登入成功，進入主畫面")

# EndRegion - Logging
# ========================================================================

# ========================================================================
# Region - 建立資料夾
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# EndRegion - 建立資料夾
# ========================================================================

# ========================================================================
# Region - 跳轉到指定 App & 截取畫面
for app_name in ux_apps:
    print(f"📸 ux_app: {app_name} 截圖完成！")
    ux_app_url_param = "#!/app/" + app_name
    
    # === 只改 hash，讓畫面重新渲染 === 
    driver.execute_script(f"window.location.hash = '{ux_app_url_param}'")
    time.sleep(DELAY)  # 等頁面query

    safe_name = app_name.replace(" ", "_")
    driver.save_screenshot(f"{SCREENSHOT_DIR}/{safe_name}.png")

# EndRegion - 跳轉到指定 App & 截取畫面
# ========================================================================

print("✅ 所有 App 截圖完成！")
driver.quit()

