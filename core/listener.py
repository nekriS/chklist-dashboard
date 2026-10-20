from system import log, loadData, saveData
import time
from addons.joke import getJoke
from core.services import get_dashboard_html
import datetime
from web.utils import get_computer_link



def getListener(dashboard):

    default_path = dashboard.config["GENERAL"]['DEFAULT_PATH']
    default_folder_path = f"{default_path}{dashboard.config["GENERAL"]['NAME_FOLDER_DATA']}"
    data_path = f"{default_folder_path}/{dashboard.config["GENERAL"]['NAME_FILE_DATA']}"

    def listener(client, event_type, context):
        log(event_type)

        try:
            user = context.get('ownerId')
            message_text = context.get('text').split('<span')[0]
            admins = [int(dashboard.config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in dashboard.config["USER_RIGHTS"]["ADMIN"].strip('[]').split(',')]

        except Exception as e:
            log(e)

        match event_type:
            case 'INCOMING_MESSAGE':

                if int(user) in admins:
                    match message_text.split(" ")[0]:
                        case "/info" | "/time":
                            info_text = f"""
Время: {datetime.datetime.now()}<br>
Версия панели: {dashboard.config["GENERAL"]["VERSION"]} 
"""
                            client.send_message(target_id=int(user), target_type="person", text=info_text)
                        case "/status":
                            pass
                        case "/dashboard" | "/db":
                            dashboard_html = get_dashboard_html(dashboard)
                            client.send_message(target_id=int(user), target_type="person", text=f"{dashboard_html}")
                        case "/joke":
                            joke = getJoke(f"Напиши короткий смешной анекдот")
                            joke = joke.replace("\n","<br>")
                            client.send_message(target_id=int(user), target_type="person", text=f"{joke}")
                        case "/getp":
                            projectId = message_text.split(" ")[1]
                            project = dashboard.projects.get_project_by_id(projectId)
                            path_last_file = project["path"].split("|")[-1]
                            client.send_file(target_id=int(user), filepath=f"{path_last_file}")
                        case "/web":
                            client.send_message(target_id=int(user), target_type="person", text=f'<a href="{get_computer_link()}">Ссылка на таблицу состояний</a>')
                        case "/complete":
                            projectId = message_text.split(" ")[1]
                            project = dashboard.projects.get_project_by_id(projectId)
                            #getProjectById(data["PROJECTS"], projectId)
                            project["state"] = 8
                            dashboard.projects.process_batch([project])
                        case "/setstate":
                            projectId = message_text.split(" ")[1]
                            project = dashboard.projects.get_project_by_id(projectId)
                            project["state"] = message_text.split(" ")[2]
                            dashboard.projects.process_batch([project])
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
