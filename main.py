from cefpython3 import cefpython as cef
import platform
import sys, ctypes

mytitle = "Test"




class MailArchiveBrowser:

    def check_versions(self):
        ver = cef.GetVersion()
        print("[hello_world.py] CEF Python {ver}".format(ver=ver["version"]))
        print("[hello_world.py] Chromium {ver}".format(ver=ver["chrome_version"]))
        print("[hello_world.py] CEF {ver}".format(ver=ver["cef_version"]))
        print("[hello_world.py] Python {ver} {arch}".format(
            ver=platform.python_version(),
            arch=platform.architecture()[0]))
        assert cef.__version__ >= "57.0", "CEF Python v57.0+ required to run this"
        

    def __init__(self):
        self.check_versions()
        sys.excepthook = cef.ExceptHook


    def change_the_window_size(self, width, height):
        window_handle = self.browser.GetOuterWindowHandle()
        SWP_NOMOVE = 0x0002
        ctypes.windll.user32.SetWindowPos(window_handle, 0,
                                          0, 0, width, height, SWP_NOMOVE)
        
        icon = ctypes.windll.user32.LoadImageW(None, "favicon.ico", ctypes.c_uint(1), 0, 0, ctypes.c_uint(0x00000010))

        # Set the window icon
        ctypes.windll.user32.SendMessageW(window_handle, 0x80, 0, icon)  # WM_SETICON message code


    def run(self):

        cef.Initialize(settings={}, switches = {"disable-web-security": ""})



        self.browser = cef.CreateBrowserSync(url="file:///resource/index.html",
                          window_title="Mail Archive Browser")
        self.change_the_window_size(1200, 600)


        self.setup_bindings()

        cef.MessageLoop()
        cef.Shutdown()

    def setup_bindings(self):
        bindings = cef.JavascriptBindings(bindToFrames=False, bindToPopups=False)
        bindings.SetFunction("printTest", self.test_function)
        bindings.SetFunction("set_mail_list", self.set_mail_list)
        self.browser.SetJavascriptBindings(bindings)


    def test_function(self, value):
        print(value)


    def set_mail_list(self):
        print("Setting Mail List")
        self.browser.ExecuteFunction("setMailList", 
                                     {"KLA Meeting": ['A', 'B', "C"], 
                                      "Lasertec Meeting": ['A1', "C1", 'D1']
                                      })


if __name__ == '__main__':
    browser = MailArchiveBrowser()
    browser.run()