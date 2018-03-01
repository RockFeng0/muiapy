# -*- encoding: utf-8 -*-
'''
Current module: muiapy.muia

Rough version history:
v1.0    Original version to use

********************************************************************
    @AUTHOR:  Administrator-Bruce Luo(罗科峰)
    MAIL:    lkf20031988@163.com
    RCS:      muiapy.muia,v 3.0 2017年4月6日
    FROM:   2016年5月9日
********************************************************************

======================================================================

UI Automation frame for IronPython.

'''


__all__ = [
    'WinWPFDriver',
    ]
    
import os,time,ctypes

DOTNET_LIB_PATH = os.path.abspath(os.path.dirname(__file__))
DOTNET_UIA_CLIENT = "UIAutomationClient.dll"
DOTNET_UIA_TYPES = "UIAutomationTypes.dll"

if os.path.isfile(os.path.join(DOTNET_LIB_PATH,DOTNET_UIA_CLIENT)) and os.path.isfile(os.path.join(DOTNET_LIB_PATH,DOTNET_UIA_TYPES)):
    import sys,clr
    sys.path.append(DOTNET_LIB_PATH)
    clr.AddReferenceToFile(DOTNET_UIA_CLIENT)
    clr.AddReferenceToFile(DOTNET_UIA_TYPES)
else:
    print "Warning: Current directory[%s] do not have .net library.\nSearching and loading dlls OK." %DOTNET_LIB_PATH

# Properties -> https://msdn.microsoft.com/zh-cn/library/system.windows.automation.automationelement(v=vs.110).aspx
PROPERTIES_LIST = ["AcceleratorKeyProperty",
    "AccessKeyProperty",
    "AutomationIdProperty",
    "BoundingRectangleProperty",
    "ClassNameProperty",
    "ControlTypeProperty",
    "ClickablePointProperty",# BoundingRectangleProperty来计算，更靠谱
    "CultureProperty",
    "FrameworkIdProperty",
    "HasKeyboardFocusProperty",
    "HelpTextProperty",
    "IsContentElementProperty",
    "IsControlElementProperty",
    "IsDockPatternAvailableProperty",
    "IsEnabledProperty",
    "IsExpandCollapsePatternAvailableProperty",
    "IsGridItemPatternAvailableProperty",
    "IsGridPatternAvailableProperty",
    "IsInvokePatternAvailableProperty",
    "IsItemContainerPatternAvailableProperty",
    "IsMultipleViewPatternAvailableProperty",
    "IsKeyboardFocusableProperty",
    "IsOffscreenProperty",
    "IsPasswordProperty",
    "IsRangeValuePatternAvailableProperty",
    "IsRequiredForFormProperty",
    "IsScrollItemPatternAvailableProperty",
    "IsScrollPatternAvailableProperty",
    "IsSelectionItemPatternAvailableProperty",
    "IsSelectionPatternAvailableProperty",
    "IsSynchronizedInputPatternAvailableProperty",
    "IsTableItemPatternAvailableProperty",
    "IsTablePatternAvailableProperty",
    "IsTextPatternAvailableProperty",
    "IsTogglePatternAvailableProperty",
    "IsTransformPatternAvailableProperty",
    "IsValuePatternAvailableProperty",
    "IsVirtualizedItemPatternAvailableProperty",
    "IsWindowPatternAvailableProperty",
    "ItemStatusProperty",
    "ItemTypeProperty",
    "LabeledByProperty",
    "LocalizedControlTypeProperty",
    "NameProperty",
    "NativeWindowHandleProperty",
    "OrientationProperty",
    "ProcessIdProperty",
    "RuntimeIdProperty"]

# Patterns -> https://msdn.microsoft.com/zh-cn/library/system.windows.automation(v=vs.110).aspx
PATTERN_LIST = ["DockPattern",
    "ExpandCollapsePattern",
    "GridItemPattern",
    "GridPattern",
    "InvokePattern",
    "ItemContainerPattern",
    "MultipleViewPattern",
    "RangeValuePattern",
    "ScrollItemPattern",
    "ScrollPattern",
    "SelectionItemPattern",
    "SelectionPattern",
    "SynchronizedInputPattern",
    "TableItemPattern",
    "TablePattern",
    "TextPattern",
    "TogglePattern",
    "TransformPattern",
    "ValuePattern",
    "VirtualizedItemPattern",
    "WindowPattern",]

# TressScope -> https://msdn.microsoft.com/zh-cn/library/system.windows.automation.treescope(v=vs.110).aspx
TREE_LIST = ["Children","Descendants"]

for m in PROPERTIES_LIST:
    exec "from System.Windows.Automation.AutomationElement import %s" %m
from System.Windows.Automation.AutomationElement import RootElement,FocusedElement,Current
from System.Windows.Automation.AutomationElement import FindAll,FindFirst,FromHandle
from System import IntPtr

from System.Windows.Automation import TreeScope
from System.Windows.Automation import ControlType
# 加载Condition派生类
from System.Windows.Automation import Condition
from System.Windows.Automation import AndCondition
from System.Windows.Automation import NotCondition
from System.Windows.Automation import OrCondition
from System.Windows.Automation import PropertyCondition
# 加载BasePattern派生类
for m in PATTERN_LIST:
    exec "from System.Windows.Automation import %s" %m

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_ulong),("y", ctypes.c_ulong)]
    
class Mouse:
    """It simulates the mouse"""
    MOUSEEVENTF_MOVE = 0x0001 # mouse move 
    MOUSEEVENTF_LEFTDOWN = 0x0002 # left button down 
    MOUSEEVENTF_LEFTUP = 0x0004 # left button up 
    MOUSEEVENTF_RIGHTDOWN = 0x0008 # right button down 
    MOUSEEVENTF_RIGHTUP = 0x0010 # right button up 
    MOUSEEVENTF_MIDDLEDOWN = 0x0020 # middle button down 
    MOUSEEVENTF_MIDDLEUP = 0x0040 # middle button up 
    MOUSEEVENTF_WHEEL = 0x0800 # wheel button rolled 
    MOUSEEVENTF_ABSOLUTE = 0x8000 # absolute move 
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    
    def __init__(self,elm):
        self.__pos = self.__get_clickable_point(elm)        
        
    def _do_event(self, flags, x_pos, y_pos, data, extra_info):
        """generate a mouse event"""
        x_calc = 65536L * x_pos / ctypes.windll.user32.GetSystemMetrics(self.SM_CXSCREEN) + 1
        y_calc = 65536L * y_pos / ctypes.windll.user32.GetSystemMetrics(self.SM_CYSCREEN) + 1
        return ctypes.windll.user32.mouse_event(flags, x_calc, y_calc, data, extra_info)

    def _get_button_value(self, button_name, button_up=False):
        """convert the name of the button into the corresponding value"""
        buttons = 0
        if button_name.find("right") >= 0:
            buttons = self.MOUSEEVENTF_RIGHTDOWN
        if button_name.find("left") >= 0:
            buttons = buttons + self.MOUSEEVENTF_LEFTDOWN
        if button_name.find("middle") >= 0:
            buttons = buttons + self.MOUSEEVENTF_MIDDLEDOWN
        if button_up:
            buttons = buttons << 1
        return buttons
    
    def __get_position(self):
        ''' return current mouse point '''
        po = POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(po))
        return int(po.x), int(po.y)         
    
    def __get_clickable_point(self, elm):
        pos = (-1, -1)
        try:
            p_obj = eval("BoundingRectangleProperty")        
            rect = elm.GetCurrentPropertyValue(p_obj)
            return int(rect.X + rect.Width / 2), int(rect.Y + rect.Height / 2)
        except:
            return pos
        
    def mouse_move(self, pos=()):
        """move the mouse to the specified coordinates"""
        if pos:
            (x, y) = pos
        else:
            (x, y) = self.__pos
        old_pos = self.__get_position()
        x =  x if (x != -1) else old_pos[0]
        y =  y if (y != -1) else old_pos[1]    
        self._do_event(self.MOUSEEVENTF_MOVE + self.MOUSEEVENTF_ABSOLUTE, x, y, 0, 0)
        print "mouse point: (%s, %s)" %(x,y)
        
    def mouse_press_button(self, pos=(), button_up=False, button_name="left"):
        """push a button of the mouse
        :param pos-> Will click elem if not pos; if defined pos, Will click the position;  pos = (600,600) etc.
        :param button_name -> left, right, middle
        :param button_up -> True or False
        """
        self.mouse_move(pos)
        self._do_event(self._get_button_value(button_name, button_up), 0, 0, 0, 0)
                    
    def mouse_click(self, pos = (), button_name= "left"):
        """Click at the specified placed
        :Usge   mouse_click((831,506), "left");mouse_click();
        :param pos-> Will click elem if not pos; if defined pos, Will click the position;  pos = (600,600) etc.
        :param button_name -> left, right, middle
        """
        self.mouse_move(pos)        
        self._do_event(self._get_button_value(button_name, False)+self._get_button_value(button_name, True), 0, 0, 0, 0)

    def mouse_double_click (self, pos = (), button_name="left"):
        """Double click at the specifed placed
        :param pos-> Will double click elem if not pos; if defined pos, Will double click the position;  pos = (600,600) etc.
        :param button_name -> left, right, middle
        """
        for i in xrange(2): 
            self.click(pos, button_name)
    
    def mouse_drag(self, spos, dpos):
        """ Drag source position to destination position
        :param spos -> source position; spos = (900,900) etc.
        :param dpos -> destination position; dpos = (600,600) etc.
        """
        if not spos or not dpos:
            return
        self.mouse_press_button(spos)
        self.mouse_press_button(dpos, button_up=True)
    
    def mouse_drag_to(self, dpos):
        """ Drag element to destination position
        :param dpos destination position; dpos = (600,600) etc.
        """
        self.mouse_press_button()
        self.mouse_press_button(dpos, button_up=True)
    
class WinWPFElement(Mouse):
    ''' Reference UISpy -> ControlPatterns '''
    
    def __init__(self,elm):
        Mouse.__init__(self, elm)
        self.__elm = elm                
    
    def getProp(self,prop):
        '''Sample usage:
            get("Name") ->NameProperty
        '''
        p = prop + "Property"
        if not p in PROPERTIES_LIST:
            return
        p_obj = eval(p)
        
        return self.__elm.GetCurrentPropertyValue(p_obj)
    
    def getControl(self,control):
        '''Sample usage:
            pattern = getControl("SelectionItem")    ->SelectionItemPattern
            print pattern.Current.IsSelected
        '''
        c = control + "Pattern"
        if not c in PATTERN_LIST:
            return        
        real_class_obj = eval(c)
        return self.__elm.GetCurrentPattern(real_class_obj.Pattern)
            
class WinWPFDriver(WinWPFElement):
    
    def __init__(self,elm=RootElement, timeout = 30):
        WinWPFElement.__init__(self, elm)        
        self.__elm = elm
        self.__timeout = timeout
    
    def find_element_by_trace(self, trace):
        ''' find elements by elements index
            Useless function, always the path is dynamic
        :trace the list of elements index.
        Sample usage:
             find_element_by_trace([0,1,2,3])
        ''' 
        elems = self.find_elements_by_children(IsControlElement = True)
        for index in trace:
            try:
                elem = elems[index]
            except Exception,e:
                print e
                return
            else:
                elems = elem.find_elements_by_children(IsControlElement = True)
        return elem
    
    def find_element_by_handle(self, hex_or_int_handle):
        ''' use dynamic handle to find element
            usage:            
                elem = find_element(AutomationId = "1",Name = "OK",ClassName = "Button")
                hex_or_int_handle = elem.getProp("NativeWindowHandle")
                elem = find_element_by_handle(hex_or_int_handle)                
        '''
        try:
            handle = hex_or_int_handle
            if isinstance(hex_or_int_handle, str):
                handle = int(hex_or_int_handle, 16);# 16进制转为 10进制;  hex(123456)->10进制转换16进制           
            
            if handle == 0:
                return
                    
            automation_element_obj = FromHandle(IntPtr(handle))        
            return WinWPFDriver(automation_element_obj)    
        except:
            return        
    
    def find_elements(self,**attr):
        ''' Need property attribute, find elements from descendants      
        Sample usage:
            elems = find_elements(AutomationId = "1",Name = "OK",ClassName = "Button")
        '''
        self.__tree = getattr(TreeScope,"Descendants")
        return self.__find_all(**attr)
    
    def find_elements_by_children(self, **attr):
        '''Need property attribute, find elements form children     --  more faster when use find_elements_by_children to branch a window tree by specification a window title.
        Sample usage:
             elems = find_elements_by_children(AutomationId = "1",Name = "OK",ClassName = "Button")
        '''        
        self.__tree = getattr(TreeScope,"Children")
        return self.__find_all(**attr)       
    
    def find_element(self, **attr):
        ''' Need property attribute, find element from descendants      
        Sample usage:
            elem = find_element(AutomationId = "1",Name = "OK",ClassName = "Button")
        '''
        self.__tree = getattr(TreeScope,"Descendants")
        return self.__find_first(**attr)
        
    def find_element_by_children(self, **attr):
        '''Need property attribute, find element form children     --  more faster when use find_element_by_children to branch a window tree by specification a window title.
        Sample usage:
             elems = find_element_by_children(AutomationId = "1",Name = "OK",ClassName = "Button")
        '''        
        self.__tree = getattr(TreeScope,"Children")
        return self.__find_first(**attr)
            
    def __until(self, method):
        end_time = time.time() + self.__timeout
        while True:
            try:
                value = method()
                if value:
                    return value
            except:
                pass
            time.sleep(0.5)
            if time.time() > end_time:
                break
        raise Exception("Timeout: not found.")
                
    def __find_all(self, **attr):
        result = []
        if not attr:
            return result
        
        conditons = []
        for k,v in attr.iteritems():
            p = k+"Property"
            if not p in PROPERTIES_LIST:
                print "Warning: can't find '%s'." %p
                return result                        
            # 转换为 属性条件
            if k == "ControlType":
                v = eval(v)
            p_obj = eval(p)
            conditons.append(PropertyCondition(p_obj, v))
            
        # Transfer to AndCondition
        if len(conditons) == 1:
            and_con = conditons[0];#必须指定一个个条件
        else:
            and_con = AndCondition(*conditons);#必须指定至少两个条件
        
        # find all elements
        elems = self.__until(lambda: self.__elm.FindAll(self.__tree, and_con))        
        for elm in elems:
            result.append(WinWPFDriver(elm, self.__timeout))
        return result
    
    def __find_first(self, **attr):        
        if not attr:
            return
        
        conditons = []
        for k,v in attr.iteritems():
            p = k+"Property"
            if not p in PROPERTIES_LIST:
                print "Warning: can't find '%s'." %p
                return                        
            # 转换为 属性条件
            if k == "ControlType":
                v = eval(v)
            p_obj = eval(p)
            conditons.append(PropertyCondition(p_obj, v))
            
        # Transfer to AndCondition
        if len(conditons) == 1:
            and_con = conditons[0];#必须指定一个个条件
        else:
            and_con = AndCondition(*conditons);#必须指定至少两个条件
        
        # find first element
        elm = self.__until(lambda: self.__elm.FindFirst(self.__tree, and_con))
        return WinWPFDriver(elm, self.__timeout)

