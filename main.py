from cefpython3 import cefpython as cef
import platform
import sys

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
        

    def run(self):
        cef.Initialize(settings={}, switches = {"disable-web-security": ""})
        self.browser = cef.CreateBrowserSync(url="file:///resource/index.html",
                          window_title="Mail Archive Browser")

        self.setup_bindings()

        cef.MessageLoop()
        cef.Shutdown()

    def setup_bindings(self):
        bindings = cef.JavascriptBindings(bindToFrames=False, bindToPopups=False)
        bindings.SetFunction("printTest", self.test_function)
        self.browser.SetJavascriptBindings(bindings)


    def test_function(self, value):
        print(value)





if __name__ == '__main__':
    browser = MailArchiveBrowser()
    browser.run()