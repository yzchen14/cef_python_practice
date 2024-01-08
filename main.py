from cefpython3 import cefpython as cef
import sys, ctypes, configparser, os
import sqlite3
from datetime import datetime

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
    def __init__(self, setting_manager):
        self.setting_manager = setting_manager
        self.db_path = self.setting_manager.setting_dict['Database']['db_path']

        if self.db_path:
            self._initialize_database()
        else:
            raise ValueError("Database path not found in settings")
        
    def _initialize_database(self):
        if not os.path.exists(self.db_path):
            print(f"Creating database at {self.db_path}")
            self._create_database()

    def _create_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE mail_archive_table (
                sender TEXT,
                entryID TEXT,
                mail_title TEXT,
                datetime TEXT,
                category TEXT,
                filepath TEXT
            )
        ''')

        conn.commit()
        conn.close()

    def add_mail_entry(self, sender, entry_id, mail_title, category, entry_datetime, filepath):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO mail_archive_table (sender, entryID, mail_title, datetime, category, filepath)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (sender, entry_id, mail_title, entry_datetime.strftime('%Y-%m-%d %H:%M:%S'), category, filepath))

        conn.commit()
        conn.close()




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
    # setting_manager = SettingManager()
    # handler = MailArchiveHandler(setting_manager)
    # custom_datetime = datetime(2023, 1, 15, 10, 30, 0)
    # handler.add_mail_entry('John Doe', '123456', 'Sample Mail', 'KLA Meetings', custom_datetime, "Test")


    browser = MailArchiveBrowser()
    browser.run()