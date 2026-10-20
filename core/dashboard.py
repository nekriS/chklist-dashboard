from messanger.link import LinkClient
from core.listener import getListener
from core.notifications import notificationsObject
from core.projects import projectsObject
from core.config import NAME_CONFIG_FILE, DEFAULT_CONFIG
import yaml
import configparser
from system import log
import os

class dashboardObject():
    def __init__(self,
                 config = configparser.ConfigParser()):
        self.config = config
        self.messanger = None
        self.projects = None
        self.notifications = None
        self.messages = None

        self._init_config()
        self.print_all_options()
        self.create_folders()
        
        self._init_messanger()
        self._init_notifications()
        self._init_projects()
        self._get_text_messages()

    def _init_config(self):
        if os.path.exists(NAME_CONFIG_FILE):
            self.update()
        else:
            for section, value in DEFAULT_CONFIG.items():
                self.config[section] = value

            with open(NAME_CONFIG_FILE, "w", encoding="utf-8") as f:
                self.config.write(f)

    def _init_projects(self):
        projects_file = f"{self.config["GENERAL"]['DEFAULT_PATH']}{self.config["GENERAL"]['NAME_FOLDER_DATA']}/projects.db"
        self.projects = projectsObject(projects_file, log=log)

    def _init_notifications(self):
        notifications_file = f"{self.config["GENERAL"]['DEFAULT_PATH']}{self.config["GENERAL"]['NAME_FOLDER_DATA']}/notifications.db"
        self.notifications = notificationsObject(notifications_file, log=log)

    def _init_messanger(self):
        bot_api_base_url = self.config["BOT"]['API_BASE_URL']
        bot_ws_url = self.config["BOT"]['WS_URL']

        self.messanger = LinkClient(bot_api_base_url, bot_ws_url, silent=False, listener=getListener(self), logger=log)
        match self.config["GENERAL"]["MODE"]:
            case "DEV1":
                self.messanger.debug = "console"
            case "DEV2":
                self.messanger.debug = "one_target"

        try:
            self.messanger.ensure_connected()
        except Exception as e:
            log(f"ERROR: {e}")

        self.messanger.download_folder = f"{self.config["GENERAL"]['DEFAULT_PATH']}{self.config["GENERAL"]['NAME_FOLDER_UPLOADS']}"
    

    def _get_text_messages(self):
        with open(f"languages/{self.config["GENERAL"]['LANGUAGE']}.yaml", 
                  "r", 
                  encoding="utf-8") as f:
            self.messages = yaml.safe_load(f)

    def update(self):
        if os.path.exists(NAME_CONFIG_FILE):
            self.config.read(NAME_CONFIG_FILE, encoding="utf-8")

    def print_all_options(self):
        log("CONFIGURATION:")
        for section, value in self.config.items():
            for param in value:
                log(f"{section} : {param} : {value[param]}")

    def create_folders(self):
        for section, value in self.config.items():
            for param in value:
                if "folder" in param:
                    dir = f"{self.config[section]['DEFAULT_PATH']}{self.config[section][param]}"
                    if not os.path.exists(dir):
                        os.makedirs(dir)
                        log(f"Directory {dir} was created!")


