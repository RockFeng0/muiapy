# 痛点（创建这个工程的目的）
我们知道，在软件自动化的过程中，常常遇到一些windows弹出框，上传文件等，常用的解决方案如AutoItv3，但它**仅仅适用于 Microsoft MFC技术的window窗口，而对于Microsoft WPF技术开发的window窗口无能为力**，创建这个项目的初衷，就是，完成 Microsoft WPF窗口的识别和操作。

* * *
# 一些基本工具

- 关于MFC窗口，常用的是auitv3的识别窗口工具   ![](https://github.com/RockFeng0/muiapy/raw/master/pic//20170421171813163.png)

- 现在我们要用这个了， UISpy 用于定位 UI，支持WPF和MFC，但是它适用于Microsoft UI 自动化(简称 MUIA) ![](https://github.com/RockFeng0/muiapy/raw/master/pic//20170421171813164.png)

# UiSpy工具的用法
	1. 确保机器上安装了 Microsoft .NET 3.0 Framework以上的版本，就像java的jdk和jre一样，这是微软的工作环境
	2. 打开UIspy，然后可以按住Ctrl键，移动鼠标到WPF的UI上，然后可以从UISpy上面的右边的Properties窗口查看AutomationElement	
	
# 效果
	我这里拿了个notepad++的安装MFC窗口，做示例
![](https://github.com/RockFeng0/muiapy/raw/master/pic//example.gif)

# 不多说，上代码，分析
```
# example.py
# encoding:utf-8

import os,time
from muia import WinWPFDriver

# 打开应用程序
os.popen(os.path.join(os.path.dirname((os.path.abspath(__file__))), "test_app", "npp.5.7.Installer.exe"))

# Use UISpy or others to spy the WPF UI.
window_title1 = u"Installer Language"
window_title2 = u"Notepad++ v5.7 安装"    
driver = WinWPFDriver()

# 寻找notepad++窗口，后面的参数，是通过UiSpy工具获取的
title1_win = driver.find_element_by_children(Name = window_title1, ControlType = "ControlType.Window")
dpos = (400,400)
# 寻找标题栏
elem = title1_win.find_element(AutomationId = u"TitleBar")
# 移动到指定位置    
elem.mouse_drag_to(dpos)    
time.sleep(1)

# 寻找English的UI 
elems = title1_win.find_elements(Name = u"English")
print "Current window element name: %s" %elems[0].getProp("Name")
# 选择英语
elems[0].getControl("SelectionItem").Select()
# 寻找Chiness的UI
elems = title1_win.find_elements(Name = u"Chinese (Simplified)")
print "Current window element name: %s" %elems[0].getProp("Name")
# 选择中文
elems[0].getControl("SelectionItem").Select()
# 寻找OK的UI
elems = title1_win.find_elements(Name = u"OK")
# 点击OK
elems[0].getControl("Invoke").Invoke()
print "---"

# 寻找安装窗口
title2_win = driver.find_element_by_children(Name = window_title2, ControlType = "ControlType.Window")
# 点击下一步
title2_win.find_elements(Name = u"下一步(N) >")[0].getControl("Invoke").Invoke()
# 点击接受
title2_win.find_elements(Name = u"我接受(I)")[0].getControl("Invoke").Invoke()
# 输入文本
title2_win.find_elements(AutomationId = "1019")[0].getControl("Value").SetValue(ur"d:\hello input")
# 点击下一步
title2_win.find_elements(Name = u"下一步(N) >")[0].getControl("Invoke").Invoke()
# 点击取消
title2_win.find_elements(Name = u"取消(C)")[0].getControl("Invoke").Invoke()

# 寻找 退出窗口
exit_win = title2_win.find_element_by_children(Name = window_title2, ControlType = "ControlType.Window")
# 点击是
exit_win.find_element(Name = u"是(Y)").mouse_click()
print "---"
```

# 整个过程，主要两个过程
1. 寻找UI		
2. 执行UI事件

# 目录结构
![](https://github.com/RockFeng0/muiapy/raw/master/pic//20170421171813165.png)

- muiapy--基于微软MUIA，我这里简单封装了一下，命名工程为muiapy
- release_version--打包成了dll，将下级目录中的包 muiapy_dll拷贝到ipy的site-packages中，就能调用 **注意使用IronPython**,如下

```
from muiapy_dll import WinWPFDriver
driver = WinWPFDriver()
```

# 微软官方，参考文档
- [在MSDN上面，关于MUIA的详细介绍](https://docs.microsoft.com/zh-cn/dotnet/framework/ui-automation/ui-automation-fundamentals)
- [关于MUIA的详细的实例](https://msdn.microsoft.com/zh-cn/magazine/dd483216.aspx)

				