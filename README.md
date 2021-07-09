# ScrapingTool
[exe版下载](https://primalg-my.sharepoint.com/:f:/g/personal/tong_primal-biz_co_jp/EhQB_6vsUwJGg79W6IJQs2oBi9UevrEapN_VLNUQ6B-MBA?e=cXDYSB)

## 已知的问题
1. resource_path在jupyter执行报错，在jupyter执行脚本时需要不启用resource_path。
2. 打包exe后chromium无法被调用。这里用了以文件夹形式打包（参数-D），打包后手动将chromium复制到打包后文件夹的方式解决。
3. 提示缺少openpyxls库，按[这里](https://blog.csdn.net/weixin_30907523/article/details/102154787)的方法解决了。

## requirements.txt
- pandas==0.24.2
- selenium==3.141.0
- chromedriver_binary==83.0.4103.39.0

## pyinstaller语句
`pyinstaller ./20210709_ScrapingScript_GreenJapan.py -D --clean -n "ScrapingTool" -i fav.ico --add-binary "./driver/chromedriver.exe;./driver" --add-binary "./browser;./browser"  `

## 参考资料
- [jupyter Notebookのコードをexe化する方法 [Anaconda3環境]](https://nprogram.hatenablog.com/entry/2019/10/21/110326)
- [Python & Selenium を PyInstaller で実行ファイル化するまと](https://www.zacoding.com/post/python-selenium-pyinstaller/)
- [【踩坑之旅】Pyinstaller的exe封装经验总结--避免文件过大和报错 - 知乎](https://zhuanlan.zhihu.com/p/144621033)
- [Tkinter选择路径功能的实现](https://blog.csdn.net/zjiang1994/article/details/53513377)
- [Python之tkinter 进度条 Progressbar](https://blog.csdn.net/qq_44168690/article/details/105092516)
- [Tkinter 教程 - 下拉列表 | D栈 - Delft Stack](https://www.delftstack.com/zh/tutorial/tkinter-tutorial/tkinter-combobox/)
- [Pandas专家总结：指定样式保存excel数据的 “N种” 姿势](https://cloud.tencent.com/developer/article/1770494)
- [Pandas 导出表格 to_execl 自动调整列宽 - 简书](https://www.jianshu.com/p/a3aed25b3c28)
