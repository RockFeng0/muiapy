# encoding:utf-8
__version__ = 3.0
try:
    from muia import WinWPFDriver
except ImportError,e:
    import sys,os,clr
    print "Warning: %s. Now Calling the dll..." %e    
    sys.path.append(os.path.abspath(os.path.dirname(__file__)))
    clr.AddReferenceToFile("muia.dll")
    clr.AddReferenceToFile("UIAutomationClient.dll")
    clr.AddReferenceToFile("UIAutomationTypes.dll")     
    from muia import WinWPFDriver