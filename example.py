#! ipy2.7
# -*- encoding: utf-8 -*-


import os
import time
from muia import WinWPFDriver

os.popen(os.path.join(os.path.dirname((os.path.abspath(__file__))), "test_app", "npp.5.7.Installer.exe"))

# Use UISpy or others to spy the WPF UI.
window_title1 = u"Installer Language"
window_title2 = u"Notepad++ v5.7 安装"
driver = WinWPFDriver()

title1_win = driver.find_element_by_children(Name=window_title1, ControlType="ControlType.Window")
dpos = (400, 400)
elem = title1_win.find_element(AutomationId=u"TitleBar")
elem.mouse_drag_to(dpos)
time.sleep(1)

elems = title1_win.find_elements(Name=u"English")
print "Current window element name: %s" % elems[0].getProp("Name")
elems[0].getControl("SelectionItem").Select()
elems = title1_win.find_elements(Name=u"Chinese (Simplified)")
print "Current window element name: %s" % elems[0].getProp("Name")
elems[0].getControl("SelectionItem").Select()
elems = title1_win.find_elements(Name=u"OK")
elems[0].getControl("Invoke").Invoke()
print "---"

title2_win = driver.find_element_by_children(Name=window_title2, ControlType="ControlType.Window")
title2_win.find_elements(Name=u"下一步(N) >")[0].getControl("Invoke").Invoke()
title2_win.find_elements(Name=u"我接受(I)")[0].getControl("Invoke").Invoke()
title2_win.find_elements(AutomationId="1019")[0].getControl("Value").SetValue(ur"d:\hello input")
title2_win.find_elements(Name=u"下一步(N) >")[0].getControl("Invoke").Invoke()
title2_win.find_elements(Name=u"取消(C)")[0].getControl("Invoke").Invoke()

exit_win = title2_win.find_element_by_children(Name=window_title2, ControlType="ControlType.Window")
exit_win.find_element(Name=u"是(Y)").mouse_click()
print "---"
