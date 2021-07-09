#!/usr/bin/env python
# coding: utf-8

# In[8]:


#制作requirements.txt
#!pipreqsnb 20210709_ScrapingScript_GreenJapan.ipynb --encoding=utf8 --force --savepath requirements.txt

#打包exe
#pyinstaller ./20210709_ScrapingScript_GreenJapan.py --onefile --clean --noconsole -n "ScrapingTool" -i fav.ico --add-binary "./driver/chromedriver.exe;./driver" --add-binary "./browser;./browser"

#（显示控制台）
#pyinstaller ./20210709_ScrapingScript_GreenJapan.py -D --clean -n "ScrapingTool" -i fav.ico --add-binary "./driver/chromedriver.exe;./driver" --add-binary "./browser;./browser"


#已知的问题
#1）resource_path在jupyter执行报错，在jupyter执行脚本时需要不启用resource_path。
#2）打包exe后chromium无法被调用。这里用了以文件夹形式打包（参数-D），打包后手动将chromium复制到打包后文件夹的方式解决。
#3）提示缺少openpyxls库，按以下方法解决：https://blog.csdn.net/weixin_30907523/article/details/102154787


# In[50]:


#GUI
import tkinter as tk 
from tkinter import END  
from tkinter.filedialog import askdirectory
from tkinter import ttk

#爬虫
import os,sys,time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import Select
import chromedriver_binary

#参考资料
#jupyter Notebookのコードをexe化する方法 [Anaconda3環境]
#https://nprogram.hatenablog.com/entry/2019/10/21/110326
#Python & Selenium を PyInstaller で実行ファイル化するまと
#https://www.zacoding.com/post/python-selenium-pyinstaller/
#PyInstaller の--add-binaryオプションを使用してブラウザとドライバも一緒にバンドルします。

#随机等待时间
import random
def timesleep(max_num=4):
    time.sleep(random.randint(2,max_num))
    
#替换文件名中不合法字符
import re
def validateTitle(title):
    rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
    new_title = re.sub(rstr, "_", title)  # 替换为下划线
    return new_title

#Pandas输出Excel自适应调整宽高
#https://cloud.tencent.com/developer/article/1770494
#https://www.jianshu.com/p/a3aed25b3c28
from openpyxl.utils import get_column_letter
from pandas import ExcelWriter
import numpy as np

def to_excel_auto_column_weight(df: pd.DataFrame, writer: ExcelWriter, sheet_name):
    """DataFrame保存为excel并自动设置列宽"""
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    #  计算表头的字符宽度
    column_widths = (
        df.columns.to_series().apply(lambda x: len(x.encode('utf-8'))).values/3*2
    )
    #  计算每列的最大字符宽度
    max_widths = (
        df.astype(str).applymap(lambda x: len(x.encode('utf-8'))).agg(max).values/3*2
    )
    # 计算整体最大宽度
    widths = np.max([column_widths, max_widths], axis=0)
    # 设置列宽
    worksheet = writer.sheets[sheet_name]
    for i, width in enumerate(widths, 1):
        # openpyxl引擎设置字符宽度时会缩水0.5左右个字符，所以干脆+2使左右都空出一个字宽。
        worksheet.column_dimensions[get_column_letter(i)].width = width + 2

#pyinstaller用 相对路径变换绝对路径
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
        #仅在 exe 运行时有效，IDE运行时报错
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


# In[13]:


#ソース内でブラウザとドライバのパスを指定する

#pyinstaller用
#browser_path = resource_path('browser/chrome.exe')
#driver_path = resource_path('driver/chromedriver.exe')

#jupyter用
browser_path='browser/chrome.exe'
driver_path='driver/chromedriver.exe'

def greenjapan(keyword):

    options = webdriver.ChromeOptions()
    options.binary_location = browser_path
    browser = webdriver.Chrome(driver_path, options=options)

    #keyword = input('検索キーワード:')
    baseurl="https://www.green-japan.com/search_key/01?keyword="+str(keyword)
    browser.get(baseurl)
    pagers=browser.find_elements_by_xpath('//div[@class="pagers"]/a')

    #计算总页数
    pagesum=1
    for i in range(len(pagers)):
        newpagenum=pagers[i].text
        try:
            if int(newpagenum) > pagesum:
                pagesum=int(newpagenum)
        except:
            pass
        
    #进度条间隔 part1（max=30）
    step=30/pagesum

    #测试用
    #pagesum=3

    #遍历搜索结果页
    data=[]
    for i in range(1,pagesum+1):
        
        #进度条（max=30）
        progressbarOne.step(step)
        window.update()
        
        url=baseurl+"&page="+str(i)
        browser.get(url)

        company_name_list=browser.find_elements_by_xpath('//h3[@class="card-info__detail-area__box__title"]') #会社名
        year_list=browser.find_elements_by_xpath('//div[@class="card-info__detail-area__box__sub-title"]/span[1]') #設立年月
        employee_list=browser.find_elements_by_xpath('//div[@class="card-info__detail-area__box__sub-title"]/span[2]') #従業員数
        age_list=browser.find_elements_by_xpath('//div[@class="card-info__detail-area__box__sub-title"]/span[3]') #平均年齢

        buz_list=browser.find_elements_by_xpath('//*[@class="card-info__detail-area__box__body"]/ul[2]/li[2]') #大業界
        smallbuz_list=browser.find_elements_by_xpath('//*[@class="card-info__detail-area__box__body"]/ul[2]/li[4]') #小業界

        hreflist=browser.find_elements_by_xpath('//*[@class="js-search-result-box card-info "]')
        timesleep()

        for cn,yr,em,ag,bz,sb,hf in zip(company_name_list,year_list,employee_list,age_list,buz_list,smallbuz_list,hreflist):
            name=cn.text
            
            ul1=yr.text
            ul2=em.text
            ul3=ag.text
            
            if "設立年月日" in ul1:
                year=ul1.replace("設立年月日 ","")
                if "従業員数" in ul2:
                    employee=ul2.replace("従業員数 ","").replace("人","")
                    if "平均年齢" in ul3:
                        age=ul3.replace("平均年齢","").replace("歳","")
                    else:
                        age=""
                else:
                    employee,age="",""
            elif "従業員数" in ul1:
                employee=ul1.replace("従業員数 ","").replace("人","")
                year=""
                if "平均年齢" in ul2:
                    age=ul2.replace("平均年齢","").replace("歳","")
                else:
                    age=""
            elif "平均年齢" in ul1:
                age=ul1.replace("平均年齢","").replace("歳","")
                year,emplyee="",""
            else:
                year,emplyee,age="","",""
                     
            employee=employee.replace("従業員数 ","").replace("人","")
            age=age.replace("平均年齢","").replace("歳","")
            
            buz=bz.text
            sbuz=sb.text
            href=hf.get_attribute("href")
            data.append([name,year,employee,age,buz,sbuz,href])

    def searchresult(content):
        #查找关键词所在位置（通过.lower()忽略大小写）
        pos=content.lower().find(keyword.lower())

        size=len(content)

        #提取关键词所在位置前后50字
        if pos==-1:
            #如果不包含关键词（返回-1）则结果为空
            res=""
        else:
            if size < 100:
                start,end=0,size
            elif pos < 50:
                start,end=0,100
            elif pos >= 50:
                if pos+50 >= size:
                    start,end=pos-50,size
                else:
                    start,end=pos-50,pos+50
            res=content[start:end].replace("\n"," ")

        return res

    #进度条间隔 part2（max=70）
    step=70/len(data)
    
    #遍历每条搜索结果，提取详情页进一步信息
    
    for n in range(len(data)):
        progressbarOne.step(step)
        window.update()
        
        joburl=data[n][6]
        browser.get(joburl)

        companyurl=browser.find_element_by_xpath('//*[@id="com_menu_com_detail"]/a').get_attribute("href") #获得公司介绍页URL
        jobcontent=browser.find_element_by_xpath('//*[@class="com_content__basic-info"]').text #获得职业介绍页全部正文

        browser.find_element_by_id("com_menu_com_detail").click()
        timesleep()
        companycontent=browser.find_element_by_xpath('//article[@class="paragraphs"]').text #获得公司介绍页全部正文

        jobtext=searchresult(jobcontent)
        comtext=searchresult(companycontent)

        data[n].append(companyurl)
        data[n].append(jobtext)
        data[n].append(comtext)

        timesleep()

    browser.close()

    df=pd.DataFrame(data,
                columns=['会社名','設立年月','従業員数',"平均年齢","業界","サブ業界",
                         "求人URL","会社説明URL","ヒット結果（求人ページ）","ヒット結果（会社説明）"])
    
    return df


# In[12]:


#Tkinter GUI

#Tkinter选择路径功能的实现 https://blog.csdn.net/zjiang1994/article/details/53513377
def selectPath():
    var_path_ = askdirectory()
    var_path.set(var_path_)
    

#https://zhuanlan.zhihu.com/p/144621033
# #------------------------------窗口-----------------------------------#
window = tk.Tk()
window.title("ScrapingTool made by tong")
window.geometry("800x600")
tk.Label(window, text="説明：", font=("MEIRYO UI", 12)).place(x=70, y=60)
tk.Label(window, text="1.検査エンジンでGreenJapanを選択し、検索キーワードを入力してから実行", font=("MEIRYO UI", 12)).place(x=70, y=90)
tk.Label(window, text="2.保存先設定をクリックして保存先を設定可能（オプション）", font=("MEIRYO UI", 12)).place(x=70, y=120)
tk.Label(window, text="検索エンジン選択", font=("MEIRYO UI", 12)).place(x=90, y=200)
tk.Label(window, text="検索キーワード入力", font=("MEIRYO UI", 12)).place(x=90, y=250)
tk.Button(window, text = "保存先設定", font=("MEIRYO UI", 12), command = selectPath).place(x = 90, y = 300)

show_text = tk.Text()
show_text.place(x=90, y=360, height=80, width=600)

#进度条 https://blog.csdn.net/qq_44168690/article/details/105092516
progressbarOne = ttk.Progressbar(window, length=600, mode='determinate', orient=tk.HORIZONTAL)
progressbarOne.place(x=90, y=460)
progressbarOne['maximum'] = 100
progressbarOne['value'] = 0

# Entry输入框，输入的值必须要定义，这里定义成字符串类型
var_site = tk.StringVar()
var_keywrod = tk.StringVar()
var_path = tk.StringVar()

#下拉菜单 https://www.delftstack.com/zh/tutorial/tkinter-tutorial/tkinter-combobox/
searchsite=ttk.Combobox(window,values=["Green Japan",
                                       "Musubu(Not Available)",
                                       "Eight(Not Available)",
                                       "PR Times（Not Available）"],
                        textvariable=var_site,
                        state="readonly") #readonly -文本字段不可编辑，用户只能从下拉列表中选择值。
searchsite.current(0) #默认为0位
searchsite.place(x=300, y=200, height=25, width=380)

# Entry输入框，输入的值必须要定义
var_keywrod = tk.Entry(window, textvariable=var_keywrod)
var_keywrod.place(x=300, y=250, height=25, width=380)

def get_tar():
    site = var_site.get()
    window_keyword = var_keywrod.get()
    filename="検索結果（キーワード："+validateTitle(window_keyword)+" データベース：GreenJapan）"
    
    if var_path.get()=="":
        #pyinstaller用
        #window_savepath=resource_path("pandas_to_excel.xlsx")
        #jupyter用
        window_savepath=filename+".xlsx"
    else:
        window_savepath = var_path.get()+"/"+filename+".xlsx"
        
    if window_keyword == "":
        show_text.insert(END,"検索キーワードを入力してください。\n")
    else:
        if site == "Green Japan":
            doc=greenjapan(window_keyword)
            #doc.to_excel(window_savepath, index=False)
            with pd.ExcelWriter(window_savepath, engine='openpyxl') as writer:
                to_excel_auto_column_weight(doc, writer, "結果")
            
            show_text.insert(END,
                             "スクレピング結果を"+window_savepath+"に保存されました。\n")
        else:
            show_text.insert(END,
                             "このサイトでのスクレピングはまだできません。\n")

get_detail = tk.Button(window, text='実行', font=("MEIRYO UI", 12), command=get_tar)
get_detail.place(x=380, y=500)

window.mainloop()


# In[ ]:




