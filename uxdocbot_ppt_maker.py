from pptx import Presentation
from pptx.util import Inches, Pt
import os
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from TM1py import TM1Service
import configparser

# ========================================================================
# Region - Definition
## Read information from config.ini (in currenct path)
config = configparser.ConfigParser()
config.read('config.ini', encoding='utf-8')

## PPT_Maker
SCREENSHOT_DIR  = config['PPT_Maker']['ppt_screenshot_path']
OUTPUT_PPTX     = config['PPT_Maker']['ppt_path']     # PPT path
UX_APP_MDX      = config['PPT_Maker']['ux_app_mdx']  # 想要製作PPT的UX Apps (define by MDX)
MASTER_PPT      = config['PPT_Maker']['master_ppt_path'] 

COVER_PAGE_IDX  = int( config['PPT_Maker']['cover_page_idx'])
FINAL_PAGE_IDX  = int( config['PPT_Maker']['final_page_idx'])
BLANK_PAGE_IDX  = int( config['PPT_Maker']['blank_page_idx'])

## UX_CS
ADDRESS         = config['UX_CS']['address']         # Content Store Address
PORT            = config['UX_CS']['port']            # Content Store Port
NAMESPACE       = config['UX_CS']['namespace']       # Content Store Namespace
SYS_USERNAME    = config['UX_CS']['system_username'] # Content Store Login System Account
SYS_PASSWORD    = config['UX_CS']['system_password'] # Content Store Login System Account Password
SSL             = config['UX_CS']['ssl']             # Content Store Using SSL?

## UX_CS Objects
dimension_name  = "}APQ UX App"
attr_name_title = "PPT Title"
attr_name_func  = "PPT Func Desc"

# EndRegion - Definition
# ========================================================================

# ========================================================================
# Region - Get UX App List
# === 連接 TM1 ===
tm1 = TM1Service(address=ADDRESS, port=PORT, user=SYS_USERNAME, password=SYS_PASSWORD, ssl=SSL, namespace=NAMESPACE)

# EndRegion - Get UX App List
# ========================================================================

# ========================================================================
# Region - Create PPT
# === 建立 PPTX ===
## PPT Template
prs = Presentation(MASTER_PPT)

# Cover Page - 封面投影片
cover_slide_layout = prs.slide_layouts[COVER_PAGE_IDX]
cover_slide = prs.slides.add_slide(cover_slide_layout)

blank_slide_layout = prs.slide_layouts[BLANK_PAGE_IDX]  # 空白投影片

# === 讀取資料夾中圖片檔案 ===
for filename in os.listdir(SCREENSHOT_DIR):
    
    name_without_ext = os.path.splitext(filename)[0]
    mdx_title = """
    SELECT
    {[}APQ UX App].[""" + name_without_ext + """]} ON ROWS,
    {[}ElementAttributes_}APQ UX App].[""" + attr_name_title + """]} ON COLUMNS
    FROM [}ElementAttributes_}APQ UX App]
    """
    cells_title = tm1.cubes.cells.execute_mdx(mdx_title)
    
    mdx_func = """
    SELECT
    {[}APQ UX App].[""" + name_without_ext + """]} ON ROWS,
    {[}ElementAttributes_}APQ UX App].[""" + attr_name_func + """]} ON COLUMNS
    FROM [}ElementAttributes_}APQ UX App]
    """
    cells_func = tm1.cubes.cells.execute_mdx(mdx_func)
    
    for coordinates, cell in cells_title.items():
        ppt_title = cell["Value"]
        
    for coordinates, cell in cells_func.items():
        func_desc = cell["Value"]

    if filename.lower().endswith(".png"):
        element = os.path.splitext(filename)[0]
        
        # 從 TM1 抓取 PPT Title attribute
        if not ppt_title:
            ppt_title = element  # 如果沒抓到，用 element 名稱當標題
            
        # 從 TM1 抓取 PPT Func Desc attribute
        if not func_desc:
            func_desc = element  # 如果沒抓到，用 element 名稱當標題
        
        # 新增投影片
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 加入標題文字
        t_left      = Inches(0.2)
        t_top       = Inches(0.15)
        t_width     = Inches(8.5)
        t_height    = Inches(0.8)
        t_box = slide.shapes.add_textbox(
                        t_left      # Title X軸
                        , t_top     # Title Y軸
                        , t_width   # Title 寬度
                        , t_height  # Title 高度
                        )
        tf = t_box.text_frame
        p = tf.add_paragraph()
        p.text = ppt_title
        p.font.size = Pt(28)
        
        
        # 插入綠色區塊（矩形）
        c_left      = Inches(0.2)
        c_top       = Inches(1.2)
        c_width     = Inches(8.5)
        c_height    = Inches(1)
        c_box = slide.shapes.add_shape(
                        MSO_AUTO_SHAPE_TYPE.RECTANGLE
                        , c_left      # Content X軸
                        , c_top       # Content Y軸
                        , c_width     # Content 寬度
                        , c_height    # Content 高度
                        )
        fill = c_box.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0, 176, 80)  # 綠色
        c_box.line.fill.background()  # 移除邊框線
        
        tf = c_box.text_frame
        tf.clear()  # 清掉預設文字
        
        # 設定文字靠左上
        tf.margin_left = Pt(5)
        tf.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP

        p = tf.paragraphs[0]
        p.text = f"Function: {func_desc}"
        p.font_size = Pt(14)
        p.font.bold = False
        p.font.color.rgb = RGBColor(255, 255, 255)  # 白字
        p.alignment = PP_ALIGN.LEFT

        # 加入圖片
        img_path = os.path.join(SCREENSHOT_DIR, filename)
        p_left      = Inches(0.2)
        p_top       = Inches(2.5)
        p_width     = Inches(9)
        p_height    = Inches(4.5)
        slide.shapes.add_picture(
            img_path
            , p_left    # 圖片X軸
            , p_top     # 圖片Y軸
            , p_width   # 圖片寬度
            , p_height  # 圖片高度
            )

# Final Page - 新增投影片
final_slide_layout = prs.slide_layouts[FINAL_PAGE_IDX]
final_slide = prs.slides.add_slide(final_slide_layout)

# === 儲存 PPTX ===


prs.save(OUTPUT_PPTX)
print(f"✅ 已成功產生 PowerPoint：{OUTPUT_PPTX}")
