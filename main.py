from cefpython3 import cefpython as cef
import sys, ctypes, configparser, os



class SettingManager:
    def __init__(self, ini_file = 'setting.ini'):
        self.ini_file = ini_file
        self.config = configparser.ConfigParser()
        self.check_or_create_settings_file()
        self.config.read(self.ini_file)
        self.setting_dict = self.get_settings()

    def check_or_create_settings_file(self):
        if not os.path.exists(self.ini_file):
            self._create_default_settings()

    def _create_default_settings(self):
        self.config['MailArchiveFolder'] = {
            'folder_path': 'E:/MailArchiveFolder',
        }
        with open(self.ini_file, 'w') as configfile:
            self.config.write(configfile)

    def get_settings(self):
        settings_dict = {}
        for section in self.config.sections():
            settings_dict[section] = dict(self.config[section])
        return settings_dict

    def update_setting(self, section, key, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, str(value))

        with open(self.ini_file, 'w') as configfile:
            self.config.write(configfile)


class MailArchiveHandler:
    pass



class MailArchiveBrowser:
    def __init__(self):
        self.setting_manager = SettingManager()
        sys.excepthook = cef.ExceptHook
        self.function_binding_dict = {
            "set_mail_list": self.set_mail_list,
            "get_setting": self.get_setting,
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


    def get_setting(self):
        print(self.setting_manager.setting_dict)
        self.browser.ExecuteFunction("setSettingValues", self.setting_manager.setting_dict)


if __name__ == '__main__':
    browser = MailArchiveBrowser()
    browser.run()