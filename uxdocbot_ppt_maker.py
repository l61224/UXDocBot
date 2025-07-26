from pptx import Presentation
from pptx.util import Inches, Pt
import os, time
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
MASTER_PPT      = config['PPT_Maker']['master_ppt_path'] 
PPT_UX_APP_MDX  = config['PPT_Maker']['ppt_ux_app_mdx']        # 想要製作投影片的的UX Apps, 將透過此UX Apps結果去尋找圖檔 (define by MDX)

COVER_PAGE_IDX  = int( config['PPT_Maker']['cover_page_idx'])
FINAL_PAGE_IDX  = int( config['PPT_Maker']['final_page_idx'])
BLANK_PAGE_IDX  = int( config['PPT_Maker']['blank_page_idx'])

TITLE_LEFT      = float( config['PPT_Maker']['title_left'])
TITLE_TOP       = float( config['PPT_Maker']['title_top'])
TITLE_WIDTH     = float( config['PPT_Maker']['title_width'])
TITLE_HEIGHT    = float( config['PPT_Maker']['title_height'])
CONTENT_LEFT    = float( config['PPT_Maker']['content_left'])
CONTENT_TOP     = float( config['PPT_Maker']['content_top'])
CONTENT_WIDTH   = float( config['PPT_Maker']['content_width'])
CONTENT_HEIGHT  = float( config['PPT_Maker']['content_height'])
PICTURE_LEFT    = float( config['PPT_Maker']['picture_left'])
PICTURE_TOP     = float( config['PPT_Maker']['picture_top'])
PICTURE_WIDTH   = float( config['PPT_Maker']['picture_width'])
PICTURE_HEIGHT  = float( config['PPT_Maker']['picture_height'])

## UX_CS
ADDRESS         = config['UX_CS']['address']         # Content Store Address
PORT            = config['UX_CS']['port']            # Content Store Port
NAMESPACE       = config['UX_CS']['namespace']       # Content Store Namespace
SYS_USERNAME    = config['UX_CS']['system_username'] # Content Store Login System Account
SYS_PASSWORD    = config['UX_CS']['system_password'] # Content Store Login System Account Password
SSL             = config['UX_CS']['ssl']             # Content Store Using SSL?

## UX_CS Objects
dimension_name  = "}APQ UX App"
attr_name_title = "Page Title"
attr_name_func  = "Page Description"
attr_name_idx   = "Page Index"
attr_name_default = "Code and Description"

# EndRegion - Definition
# ========================================================================

# ========================================================================
# Region - Get UX App List
# === 連接 TM1 ===
tm1 = TM1Service(address=ADDRESS, port=PORT, user=SYS_USERNAME, password=SYS_PASSWORD, ssl=SSL, namespace=NAMESPACE)
mdx = 'Order( ' + PPT_UX_APP_MDX + ' , [' + dimension_name + '].[' + dimension_name + '].CurrentMember.Properties("' + attr_name_idx + '"), ASC)'
ux_apps = tm1.elements.execute_set_mdx_element_names( mdx)
 
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

# Blank Page - 空白投影片
blank_slide_layout = prs.slide_layouts[BLANK_PAGE_IDX]

# EndRegion - Create PPT
# ========================================================================

# ========================================================================
# Region - Create Page Title/ Page Description dictionary
title_dict = tm1.elements.get_attribute_of_elements(
                dimension_name='}APQ UX App', 
                hierarchy_name='}APQ UX App',
                attribute=attr_name_title,
            )
            
desc_dict = tm1.elements.get_attribute_of_elements(
                dimension_name='}APQ UX App', 
                hierarchy_name='}APQ UX App',
                attribute=attr_name_func,
            )

code_desc_dict = tm1.elements.get_attribute_of_elements(
                dimension_name='}APQ UX App', 
                hierarchy_name='}APQ UX App',
                attribute=attr_name_default,
            )
# EndRegion - Create Page Title/ Page Description dictionary
# ========================================================================

# ========================================================================
# Region - Insert Slides by screenshot            
# === Loop UX App element by sorting ===
for app_name in ux_apps:
    filename = app_name + '.png'
    if app_name:
        element = app_name
        # 從 TM1 抓取 Page Title attribute
        ppt_title = title_dict.get(element, element)
        if ppt_title == element:
            ppt_title = code_desc_dict.get(element, element)  # 如果沒抓到，用 Code and Description attribute當標題
            
        # 從 TM1 抓取 Page Description attribute
        func_desc = desc_dict.get(element, element)
        if not func_desc:
            func_desc = element  # 如果沒抓到，用 element 名稱當標題
        
        # 新增Blank投影片
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 加入標題文字
        t_left      = Inches(TITLE_LEFT)
        t_top       = Inches(TITLE_TOP)
        t_width     = Inches(TITLE_WIDTH)
        t_height    = Inches(TITLE_HEIGHT)
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
        c_left      = Inches(CONTENT_LEFT)
        c_top       = Inches(CONTENT_TOP)
        c_width     = Inches(CONTENT_WIDTH)
        c_height    = Inches(CONTENT_HEIGHT)
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
        img_path    = os.path.join(SCREENSHOT_DIR, filename)
        p_left      = Inches(PICTURE_LEFT)
        p_top       = Inches(PICTURE_TOP)
        p_width     = Inches(PICTURE_WIDTH)
        p_height    = Inches(PICTURE_HEIGHT)
        slide.shapes.add_picture(
            img_path
            , p_left    # 圖片X軸
            , p_top     # 圖片Y軸
            , p_width   # 圖片寬度
            , p_height  # 圖片高度
            )

# EndRegion - Insert Slides by screenshot
# ========================================================================

# ========================================================================
# Region - Final Page
# Final Page
final_slide_layout = prs.slide_layouts[FINAL_PAGE_IDX]
final_slide = prs.slides.add_slide(final_slide_layout)

# EndRegion - Final Page
# ========================================================================

# === 儲存 PPTX ===
prs.save(OUTPUT_PPTX)
print(f"✅ 已成功產生 PowerPoint：{OUTPUT_PPTX}")
time.sleep(3)
