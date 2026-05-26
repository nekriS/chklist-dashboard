from system import log, loadData, saveData
import time
from addons.joke import getJoke
from core.services import get_dashboard_html
import datetime
from web.utils import get_computer_link

def getProjectById(projects, id):
    for project in projects.keys():
        if not(project == "_COUNTER"):
            if int(projects[project]["NUMBER"]) == int(id):
                return projects[project]

def getListener(config):

    def listener(client, event_type, context):
        log(event_type)
        #print(context)
        try:
            user = context.get('ownerId')
            message_text = context.get('text').split('<span')[0]

            admins = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"]["ADMIN"].strip('[]').split(',')]
            #aliase = config["USER_ALIASE"]
        except:
            pass
        #admins = self.config["USER_RIGHTS"]["ADMIN"]
        data = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
        match event_type:
            case 'INCOMING_MESSAGE':

                if int(user) in admins:
                    match message_text.split(" ")[0]:
                        case "/info" | "/time":
                            client.send_message(target_id=int(user), target_type="person", text=f"Время: {datetime.datetime.now()}")
                        case "/status":
                            pass
                        case "/dashboard" | "/db":
                            dashboard_html = get_dashboard_html(config)
                            client.send_message(target_id=int(user), target_type="person", text=f"{dashboard_html}")
                        case "/joke":
                            joke = getJoke(f"Напиши короткий смешной анекдот")
                            joke = joke.replace("\n","<br>")
                            client.send_message(target_id=int(user), target_type="person", text=f"{joke}")
                        case "/getp":
                            projectId = message_text.split(" ")[1]
                            projects = data["PROJECTS"]
                            path_last_file = getProjectById(projects, projectId)["PATH"][-1]
                            client.send_file(target_id=int(user), filepath=f"{path_last_file}")
                        case "/web":
                            client.send_message(target_id=int(user), target_type="person", text=f'<a href="{get_computer_link()}">Ссылка на таблицу состояний</a>')
                        case "/complete":
                            projectId = message_text.split(" ")[1]
                            project = getProjectById(data["PROJECTS"], projectId)
                            project["STATE"] = 8
                            saveData(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
                        case "/setstate":
                            projectId = message_text.split(" ")[1]
                            project = getProjectById(data["PROJECTS"], projectId)
                            project["STATE"] = message_text.split(" ")[2]
                            saveData(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
                        case "/help":
                            help_message = """
<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;">
<strong>Список доступных команд:</strong><br><br>
<strong>/help</strong> - Вызов справки.<br>
<strong>/web</strong> - Получить ссылку на web версию таблицы состояний.<br>
<strong>/getp <id></strong> - Получить актуальный чеклист по проекту с нужным id.<br>
<strong>/dashboard</strong> - Получить таблицу состояний.<br>
<strong>/info</strong> - Получить информацию о хосте.<br>
<strong>/setstate <id> <state></strong> - Установить состояние проекта.<br>
<strong>/complete <id></strong> - Завершить проект вручную.<br>
</div>
"""
                            client.send_message(target_id=int(user), target_type="person", text=help_message)

                    log(f"admin {user}")
            case "CONNECTION_CLOSED":
                log("Connection is closed! Try again...")
                try:
                    client.ensure_connected()
                    while not client.is_connected:
                        pass
                    else:
                        client.subscribe("/topic/messages")
                except:
                    time.sleep(30)
                    log("Connection is aborted! Try again against 30 seconds!")
                    listener(client, event_type, context)
            
            case "FILE_RECEIVED":
                client.send_message(target_id=int(user), target_type="person", text=f"Файл получен!")

    return listener
