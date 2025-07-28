<img width="460" height="330" alt="DocBot_w_BK" src="https://github.com/user-attachments/assets/ffd62e5d-042d-4e7f-8fb2-7d137ee31272" />

# ❓Why UXDocBot🤖
UX page changes frequently, requiring frequent and time-consuming screenshots. Capturing screenshots based on user permissions adds even more time.  
UXDocBot automates screenshots📸 and PowerPoint creation📋 to save your time and standardize documentation.

- **📸Screen Catcher**
    - One-click automated screenshots of your ux page
    - Capture screen content based on user permissions
- **📋PPT Maker** 
    - Create slides with custom order
    - Easily edit titles and contents of each slide to fit your needs 
  
# 📒Installation
#### Python:
>
>pip install -r requirements.txt
>

#### Content Store:
1. Copy all objects from the `main/` folder into your content store instance data directory.
2. Restart content store instance.
3. Run TI: 'z.UXDocBot_V1.0'

 # ✏️Configuration
 1. Adjust `config.ini` content
#### `[UX_CS]`
```
  Content store instance information & Login user information
```
#### `[Screen_Catcher]`
```
  ux_base_url:      Base URL for UX you want to snapshot
  screenshot_path:  Screenshot storage folder
  ux_app_mdx:       UX App MDX those pages you want to do snapshot
```
#### `[PPT_Maker]`
```
  ppt_screenshot_path:    Screenshot storage folder
  master_ppt_path:        Master Slides path
  ppt_path:               The generated PPT storage path
  ppt_ux_app_mdx:         UX App MDX those pages you want to do PPT (PPT_Maker will use this mdx result to find the screenshot in ppt_screenshot_path
```
 # 💪Start your first UXDocBot journey
 1. Screen Catcher:
    1) Execute `uxdocbot_screen_catcher.py` after you completed the configuration of `config.ini`.
    2) Screen Catcher will take screenshots according to the app ID you specified by mdx and save them to the specified path.
 2. PPT Maker:
    1) Maintain the following information in cube: }ElementAttributes_}APQ UX App
       1) Page Title
       2) Page Description
       3) Page Index
    2) Execute `uxdocbot_ppt_maker.py` then PPT Maker will generate a PPT according to the content and order you maintain.

 # 🙏 Community & Contribution
This project is an ongoing effort to simplify the process of documenting **Apliqo UX** screens and help the TM1 community automate UI screenshot and PowerPoint generation workflows.

If you find this useful or have ideas to make it better, please feel free to:
* Open issues to report bugs or request features
* Submit pull requests with improvements or fixes
* Share your feedback and use cases to help evolve this tool

This project is an ongoing effort to simplify the process of documenting and help the TM1 community automate UI screenshot and PowerPoint generation workflows. 🚀

# LICENSE
This project is licensed under the MIT License - see the [LICENSE](https://github.com/l61224/UXDocBot/blob/main/LICENSE) file for details.
