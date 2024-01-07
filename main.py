from cefpython3 import cefpython as cef
import platform
import sys, ctypes


class MailArchiveBrowser:
    def __init__(self):
        sys.excepthook = cef.ExceptHook
        self.function_binding_dict = {
            "set_mail_list": self.set_mail_list
        }

    def change_the_window_size_and_icon(self, width, height):
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
        self.change_the_window_size_and_icon(1200, 600)
        self.setup_bindings()
        cef.MessageLoop()
        cef.Shutdown()

    def setup_bindings(self):
        bindings = cef.JavascriptBindings(bindToFrames=False, bindToPopups=False)
        for key in self.function_binding_dict:
            bindings.SetFunction(key, self.function_binding_dict[key])
        self.browser.SetJavascriptBindings(bindings)

    def set_mail_list(self):
        print("Setting Mail List")
        self.browser.ExecuteFunction("setMailList", 
                                     {"KLA Meeting": ['A', 'B', "C"], 
                                      "Lasertec Meeting": ['A1', "C1", 'D1']
                                      })


if __name__ == '__main__':
    browser = MailArchiveBrowser()
    browser.run()