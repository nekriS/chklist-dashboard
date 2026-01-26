import datetime
import pandas as pd
from system import load_data, save_data
from pathlib import Path
from pyodbc import connect
from system import log
import shutil
#from htmls import get_dashboard_html
from joke import get_joke
from excel import get_column_letter, add_outer_border, set_column_autowidth, create_head, thin_border_, Font, Alignment, oxl, PatternFill
import time
import os




def get_components(config, data, project):
    components = {}
    try:
        TMP_PARTS = pd.read_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD']}").dropna(subset=["NAME"])['TMP_NAME'].values
    except:
        TMP_PARTS = []

    for path in data[project]["PATH"]:

        with pd.option_context('future.no_silent_downcasting', True):
            df = pd.read_excel(f"{path}", header=None, names=["B", "C", "D", "E", "F", "J", "K"], usecols="B, C, D, E, F, J, K",skiprows=7).replace(to_replace={'Не проверено': 0, 'Да': 1, 'Нет': -1})
            column_order = ["B", "C", "D", "E", "F", "J", "K"]
            df = df[column_order].fillna('_NONE')
            table = df.to_numpy()

        for lines in table:

            if lines[1] == -1 or lines[3] == -1:
                lines13 = -1
            elif lines[1] == 1 and lines[3] == 1:
                lines13 = 1
            else:
                lines13 = 0

            if lines[0] not in components:
                components[lines[0]] = [lines13, lines[5], lines[2], lines[4], lines[6]]
            else:
                if lines13 != 0:
                    components[lines[0]][0] = lines13
                if lines[5] != 0:
                    components[lines[0]][1] = lines[5]
                if lines[2] != "_NONE":
                    components[lines[0]][2] = lines[2]
                if lines[4] != "_NONE":
                    components[lines[0]][3] = lines[4]
                if lines[6] != "_NONE":
                    components[lines[0]][4] = lines[6]

    for component in components.keys():
        if component in TMP_PARTS:
            components[component].append(1)
        else:
            components[component].append(0)

    return components

def send_last_file(client, config, target, text, file):

    if isinstance(target, int):
        targets = [target]
    elif isinstance(target, str):
        targets = [int(config["USER_ALIASE"][x.upper().strip()]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]

    for id in targets:
        client.send_file(target_id = id, filepath = file, target_type="person", message=text)


def get_list_ckeckers(components):

    # {'TMP-1487': [1, 0, 'CAP_FP', 'GME', '_NONE', 0],
    sch_checkers = {}
    pcb_checkers = {}
    for key in components.keys():
        component = components[key]
        
        if component[3] != "_NONE":
            
            if component[3] not in sch_checkers.keys():
                sch_checkers[component[3]] = []

            # Добавляем проверяющего в список, только если комопнент не проверен и не добавлен в базу
            if component[0] == 0 and component[5] != 1:
                sch_checkers[component[3]].append(key)
        else:
            return False
        
        if component[4] != "_NONE":

            if component[4] not in pcb_checkers.keys():
                pcb_checkers[component[4]] = []

            if component[1] == 0 and component[5] != 1:
                pcb_checkers[component[4]].append(key)
        else:
            return False
    
    checkers = {}
    checkers["sch"] = sch_checkers
    checkers["pcb"] = pcb_checkers

    return checkers





def check_status(client, config):

    data = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    projects = data["PROJECTS"]

    for project in projects:
        if not(project == "_COUNTER"):

            state = projects[project]["STATE"]
            path_last_file = projects[project]["PATH"][-1]
            developer = projects[project]["DEVELOPER"]
            components = get_components(config, data, project)

            match state:
                case 0: # Новый компонент

                    message = f"""
Добавлен новый проект для проверки!<br>
<br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество компонентов: {len(components)}
"""
                    send_last_file(client, config, "lib_manager", message, path_last_file)
                    create_task_notification(config, data, project, "lib_manager", 0, step=7)
                    projects[project]["STATE"] = 1
                
                case 1: # Компонент отослан библиотекарю на распределение

                    checkers = get_list_ckeckers(components)
                    if checkers:
                        delete_task_notification(config, data, project, "lib_manager", 0)
                        projects[project]["STATE"] = 2

                case 2: # Получен файл от библиотекаря

                    checkers = get_list_ckeckers(components)
                    if checkers:
                        for checker in checkers["sch"].keys():
                            
                            if len(checkers["sch"][checker]) > 0:

                                message = f"""
Вы были назначены библиотекарем схем символов компонентов!<br>
<br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["sch"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                send_last_file(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 1, step=7)

                        for checker in checkers["pcb"].keys():
                            if len(checkers["pcb"][checker]) > 0:

                                message = f"""
Вы были назначены библиотекарем посадочных компонентов!<br>
<br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["sch"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                send_last_file(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 1, step=7)
                        
                        projects[project]["STATE"] = 3

                    else:
                        projects[project]["STATE"] = 1

                case 3: # Файлы разосланы исполнителям, но никто не отослал файл
                    pass

                case 4: # Идет проверка, кто-то прислал файл
                    pass 

                case 5: # Все проверено
                    pass

                case 9: # Готово номинально, если все компоненты уже в базе
                    pass
                
                case 10: # Готово
                    pass




                    
                
            

# def update_status(config):
#     data_all = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
#     data = data_all["PROJECTS"]
#     for project in data:
#         if not(project == "_COUNTER"):

#             components = get_components(config, data, project)
#             yes_bd_schem = 0
#             no_bd_schem = 0
#             pcb_yes = 0
#             noTMP = 0
            
#             for comp in components:
#                 if components[comp][0] == 1:
#                     yes_bd_schem += 1
#                 elif components[comp][0] == -1:
#                     no_bd_schem -= 1
#                 if components[comp][1] == 1:
#                     pcb_yes += 1
#                 noTMP += (components[comp][5])


#             if data[project]["BD"] != yes_bd_schem:
#                 data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#             data[project]["BD"] = yes_bd_schem

#             if data[project]["FP"] != pcb_yes:
#                 data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#             data[project]["FP"] = pcb_yes

#             if data[project]["COUNT"] != len(components):
#                 data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#             data[project]["COUNT"] = len(components)

#             if data[project]["noTMP"] != noTMP:
#                 data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#             data[project]["noTMP"] = noTMP

#             if yes_bd_schem == data[project]["COUNT"] and pcb_yes == data[project]["COUNT"] and noTMP == data[project]["COUNT"]:
#                 if data[project]["STATUS"] != "Готов":
#                     data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#                 data[project]["STATUS"] = "Готов"
#             elif (noTMP == data[project]["COUNT"]):
#                 if data[project]["STATUS"] != "Готов номинально":
#                     data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#                 data[project]["STATUS"] = "Готов номинально"
#             elif yes_bd_schem == data[project]["COUNT"] and pcb_yes == data[project]["COUNT"]:
#                 if data[project]["STATUS"] != "К переводу":
#                     data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#                 data[project]["STATUS"] = "К переводу"
#             elif ((yes_bd_schem < data[project]["COUNT"] or pcb_yes < data[project]["COUNT"]) and (yes_bd_schem > 0 or pcb_yes > 0)) or (no_bd_schem > 0):
#                 if data[project]["STATUS"] != "Идет проверка":
#                     data[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#                 data[project]["STATUS"] = "Идет проверка"
            
#     data_all["PROJECTS"] = data
#     save_data(data_all, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")


def get_dataframe(data):

    data_table = []

    for project in data:

        if not(project == "_COUNTER"):

            if not data[project]["VISIBLE"]:
                continue
            
            row_data = []
            row_data.append(data[project]["NUMBER"])
            row_data.append(data[project]["PROJECT_NAME"])
            row_data.append(data[project]["DEVELOPER"])
            row_data.append(data[project]["STATUS"])
            row_data.append(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"))
            row_data.append(f"{data[project]["BD"]:>3} / {data[project]["COUNT"]:<3}")
            row_data.append(f"{data[project]["FP"]:>3} / {data[project]["COUNT"]:<3}")
            row_data.append(f"{data[project]["noTMP"]:>3} / {data[project]["COUNT"]:<3}")

            data_table.append(row_data)

    data_df = pd.DataFrame(data_table)
    #print(data_df)
    status_order = ["Новый", "Идет проверка", "К переводу", "Готов", "Готов номинально", "Готов вручную"][::-1]
    #print(status_order)
    data_df[3] = pd.Categorical(data_df[3], categories=status_order, ordered=True)
    data_df_s = data_df.sort_values(3).reset_index(drop=True)

    data_df_s.columns = ['ID', 'Имя проекта', 'Разработчик', 'Статус', 'Последнее обновление', 'БД+Символ', 'Посадочное', 'Переведен в постоянные']
    return data_df_s


def get_tmp_parts_from_db(config):

    try:
        #'SQL_CONNECTION_STRING': "Driver={SQL Server};Server=CADENCESTC\SQLEXPRESS;Database=CIP_E;Uid=CIP_E_CIS_User;Pwd=Test1234;"
        DRIVER = "{" + config["DB"]["DRIVER"] + "}"
        SQL_CONNECTION_STRING = f'Driver={DRIVER};Server={config["DB"]["SERVER"]};Database={config["DB"]["DATABASE"]};Uid={config["DB"]["USER"]};Pwd={config["DB"]["PASSWORD"]};'
        conn = connect(SQL_CONNECTION_STRING)

        SQL_QUERY = """
        SELECT *
        FROM TMPPRTS;
        """

        cursor = conn.cursor()
        cursor.execute(SQL_QUERY)
        records = cursor.fetchall()
        table = []
        for r in records:
            line = []
            for m in r:
                line.append(m)
            table.append(line)

        df = pd.DataFrame(table, index=None, columns=["TMP_NAME", "NAME", "AUTHOR", "DATE"])
        #df = df.dropna(subset=["NAME"])
        df.to_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD']}", index=False)
        log(f"BD with TMP parts was updated successful.")

    except Exception as e:
        log(f"ERROR: {e}")

def add_new_checklist(data, path):
    try:
        #data = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
        file_head = pd.read_excel(f"{path}", header=None, nrows=4, usecols="B").fillna("_NONE").values
        table = pd.read_excel(f"{path}", header=None, usecols="A",skiprows=7).fillna("_NONE").values
        project_name = file_head[0][0]
        
        checklist = {}
        checklist["NUMBER"] = data["_COUNTER"]
        data["_COUNTER"] += 1
        checklist["PROJECT_NAME"] = project_name
        checklist["PATH"] = [path]
        checklist["DEVELOPER"] = file_head[1][0]
        #checklist["STATUS"] = "Новый"
        checklist["STATE"] = 0
        checklist["COUNT"] = len(table)
        checklist["BD"] = 0
        checklist["FP"] = 0
        checklist["noTMP"] = 0
        checklist["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        checklist["VISIBLE"] = True 

        data[project_name] = checklist

        #save_data(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
        log(f"Project {project_name} was added!")
    except Exception as e:
        log(f"ERROR: {e}")
    #update_dashboard(config)


def check_date_diff(date1, date2):
    diff = (date2 - date1).days
    return diff

def delete_project(project):
    pass


def check_function(config, client):

    data_all = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    data = data_all["PROJECTS"]
    path_upload_folder = f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_UPLOADS']}"
    path_work_folder = f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_INWORK']}"
    today_date = datetime.datetime.now().strftime('%Y_%m_%d')
    current_time = datetime.datetime.now().strftime('%H_%M_%S')
    upload_files = os.listdir(path_upload_folder)
    if len(upload_files) > 0:
        for file in upload_files:
            try:
                pass
                file_head = pd.read_excel(f"{path_upload_folder}/{file}", header=None, nrows=4, usecols="B").fillna("_NONE").values
                project_name = file_head[0][0]
                new_file_path = f"{path_work_folder}/{project_name}_{today_date}_{current_time}.xlsx".replace(" ", "_")
                shutil.move(Path(f"{path_upload_folder}/{file}"), Path(new_file_path))

                if not(project_name in data):
                    add_new_checklist(data, new_file_path)
                else:
                    data[project_name]["PATH"].append(new_file_path)
                    data[project_name]["COUNT"] = max(data[project_name]["COUNT"], len(pd.read_excel(f"{new_file_path}", header=None, usecols="A",skiprows=7).fillna("_NONE").values))

                    client.send_message()
                    
                    log(f"New PATH in {project_name} was added!")
            except:
                pass
        
        #update_dashboard(config)
    
    for project in data:
        if not(project == "_COUNTER"):
            if "Готов" in data[project]["STATUS"]:
                if check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_INVISIBLE"]) and bool(data[project]["VISIBLE"]) == True:
                    data[project]["VISIBLE"] = False
                    log(f"Project {project} has become invisible!")
                if check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_DELETE"]) and bool(data[project]["VISIBLE"]) == False:
                    delete_project(project)

            if data[project]["noTMP"] < data[project]["COUNT"] and data[project]["BD"] == data[project]["COUNT"] and data[project]["FP"] == data[project]["COUNT"]:
                data_all = create_task_notification(config, data_all, project, "lib_manager", 0)
            elif (data[project]["noTMP"] == data[project]["COUNT"] and data[project]["BD"] == data[project]["COUNT"] and data[project]["FP"] == data[project]["COUNT"]) or data[project]["noTMP"] == 0:
                data_all = delete_task_notification(config, data_all, project, "lib_manager", 0)

    
    data_all["NOTIFICATIONS"] = send_notifications(client, data_all["NOTIFICATIONS"])

    data_all["PROJECTS"] = data
    save_data(data_all, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    #update_status(config)
    try:
        update_dashboard(config)
    except Exception as e:
        log(f"[WARN] Failed to update dashboard: {e}")

def create_task_notification(config, data_all, project, target, type_n, step=3, count=356):
    """
    Docstring для create_task_notification
    
    :param config: Конфигурация
    :param data_all: Данные
    :param project: Проект
    :param target: Цель, конкретный пользователь или группа
    :param type_n: Тип
    :param step: Шаг в днях (по умолчанию 3)
    :param count: Количество (по умолчанию 356)
    """

    #data["NOTIFICATION"] = []
    if isinstance(target, int):
        targets = [target]
    elif isinstance(target, str):
        targets = [int(config["USER_ALIASE"][x.upper().strip()]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
        #targets = []
    else:
        return data_all
    for target in targets:
        if f"{project}-{target}-{type_n}" not in list(data_all["NOTIFICATIONS"].keys()):
            notification = {}
            notification["PROJECT_NAME"] = project
            notification["TARGET"] = target
            notification["STEP"] = step
            notification["NEXT_TIME"] = f"{datetime.datetime.now().strftime('%Y.%m.%d')}"
            notification["COUNT"] = count
            notification["TYPE"] = type_n
            data_all["NOTIFICATIONS"][f"{project}-{target}-{type_n}"] = notification
            log(f"Notification {project}-{target}-{type_n} was created!")
    return data_all

def delete_task_notification(config, data_all, project, target, type_n):

    #data["NOTIFICATION"] = []
    if isinstance(target, int):
        targets = [target]
    elif isinstance(target, str):
        targets = [int(config["USER_ALIASE"][x.upper().strip()]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
        #targets = []
    else:
        return data_all
    
    for target in targets:
        if f"{project}-{target}-{type_n}" in list(data_all["NOTIFICATIONS"].keys()):
            del data_all["NOTIFICATIONS"][f"{project}-{target}-{type_n}"]
            log(f"Notification {project}-{target}-{type_n} was deleted!")

    return data_all



def send_notifications(client, notifications):

    for notice in notifications.keys():

        if abs(check_date_diff(datetime.datetime.strptime(notifications[notice]["NEXT_TIME"], "%Y.%m.%d"), datetime.datetime.now())) < 1:

            notifications[notice]
            match notifications[notice]["TYPE"]:
                case 0:
                    text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>В проекте <strong>{notifications[notice]["PROJECT_NAME"]}</strong> все компоненты готовы к переводу в базу!</div>'
                    #text = f"<strong>Внимание!</strong><br><br>В проекте <mark>{notifications[notice]["PROJECT_NAME"]}</mark> все компоненты готовы к переводу в базу!"
                    #f"<strong>Внимание!</strong><br><br>В проекте <mark>CUBESAT_TRANSCEIVER_2_25033_R1</mark> все компоненты готовы к переводу в базу!"
                case _:
                    text = None
                
            if text is not None:
                client.send_message(target_id=int(notifications[notice]["TARGET"]), target_type="person", text=text)
                notifications[notice]["COUNT"] -= 1
                if notifications[notice]["COUNT"] > 0:
                    #notifications[notice]["NEXT_TIME"] = notifications[notice]["NEXT_TIME"]
                    #notifications[notice]["NEXT_TIME"]
                    notifications[notice]["NEXT_TIME"] = (datetime.datetime.strptime(notifications[notice]["NEXT_TIME"], "%Y.%m.%d") + datetime.timedelta(days=notifications[notice]["STEP"])).strftime('%Y.%m.%d')
                else:
                    del notifications[notice]
                    #check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now())

    return notifications



def start_event_checker(config):

    def event_checker(client, event_type, context):
        log(event_type)
        #print(context)
        try:
            user = context.get('ownerId')
            message_text = context.get('text').split('<span')[0]

            admins = [int(config["USER_ALIASE"][x.upper().strip()]) for x in config["USER_RIGHTS"]["ADMIN"].strip('[]').split(',')]
            #aliase = config["USER_ALIASE"]
        except:
            pass
        #admins = self.config["USER_RIGHTS"]["ADMIN"]
        match event_type:
            case 'INCOMING_MESSAGE':

                if int(user) in admins:
                    match message_text.split(" ")[0]:
                        case "/time":
                            client.send_message(target_id=int(user), target_type="person", text=f"Время: {datetime.datetime.now()}")
                        case "/status":
                            pass
                        case "/dashboard":
                            dashboard_html = get_dashboard_html(config)
                            client.send_message(target_id=int(user), target_type="person", text=f"{dashboard_html}")
                        case "/joke":
                            joke = get_joke(f"Напиши короткий смешной анекдот про корпоративный мессенджер Postlink.")
                            joke = joke.replace("\n","<br>")
                            client.send_message(target_id=int(user), target_type="person", text=f"{joke}")

                            

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
                    event_checker(client, event_type, context)

    return event_checker


def update_dashboard(config):

    path_dashboard_file = f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FILE_DASHBOARD']}"
    today_date = datetime.datetime.now().strftime('%Y-%m-%d')
    current_time = datetime.datetime.now().strftime('%H:%M:%S')

    wb = oxl.Workbook()
    ws = wb.active
    ws.title = "Сводная таблица"
    create_head(ws)

    ws['B2'] = datetime.datetime.now()
    ws['B2'].alignment = Alignment(horizontal='left')

    data_all = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    data = data_all["PROJECTS"]
    row_id = 4
    
    data_df_s = get_dataframe(data)
    #print(data_df_s)
    
    for row_data in data_df_s.values:

            project = row_data[1]

            # if not data[project]["VISIBLE"]:
            #     continue

            components = get_components(config, data, project)
            #print(components)

            # row_data = []
            # row_data.append(data[project]["NUMBER"])
            # row_data.append(data[project]["PROJECT_NAME"])
            # row_data.append(data[project]["DEVELOPER"])
            # row_data.append(data[project]["STATUS"])
            # row_data.append(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"))
            # row_data.append(f"{data[project]["BD"]:>3} / {data[project]["COUNT"]:<3}")
            # row_data.append(f"{data[project]["FP"]:>3} / {data[project]["COUNT"]:<3}")
            # row_data.append(f"{data[project]["noTMP"]:>3} / {data[project]["COUNT"]:<3}")


            match data[project]["STATUS"]:
                case "Новый":
                    color = PatternFill(start_color=config['COLORS']['YELLOW'], end_color=config['COLORS']['YELLOW'], fill_type="solid")
                case "Идет проверка":
                    color = PatternFill(start_color=config['COLORS']['RED'], end_color=config['COLORS']['RED'], fill_type="solid")
                case "К переводу":
                    color = PatternFill(start_color=config['COLORS']['BLUE'], end_color=config['COLORS']['BLUE'], fill_type="solid")
                case "Готов":
                    color = PatternFill(start_color=config['COLORS']['GREEN'], end_color=config['COLORS']['GREEN'], fill_type="solid")
                case "Готов номинально":
                    color = PatternFill(start_color=config['COLORS']['GREEN'], end_color=config['COLORS']['GREEN'], fill_type="solid")
                case "Готов вручную":
                    color = PatternFill(start_color=config['COLORS']['GREEN'], end_color=config['COLORS']['GREEN'], fill_type="solid")

            for col_idx, value in enumerate(row_data, start=1):

                style_center = Alignment(horizontal='center', vertical='center')
                ws.cell(row=row_id, column=col_idx, value=value).alignment = style_center
                ws.cell(row=row_id, column=col_idx).fill = color
                ws.cell(row=row_id, column=col_idx).border = thin_border_
                ws.row_dimensions[row_id].height = 21

                if col_idx == 2:
                    new_sheet = wb.create_sheet(title=value)

                    hyperlink_cell1 = new_sheet.cell(row=1, column=1)  
                    hyperlink_cell1.value = "К сводной таблице"
                    hyperlink_cell1.hyperlink = f"#'{ws.title}'!{get_column_letter(col_idx)}{row_id}" 
                    hyperlink_cell1.font = Font(underline="single", color="0563C1")  
                    #new_sheet.cell(row=1, column=1).alignment = alignment_center

                    new_sheet['A2'] = value
                    new_sheet['A3'] = "Компонент"
                    new_sheet['B3'] = "Схемсимвол"
                    new_sheet['C3'] = "БД+Символ"
                    #new_sheet['C3'] = "БД+Символ Нет"
                    new_sheet['D3'] = "Посадочное"
                    new_sheet['E3'] = "Перевод"

                    style_center = Alignment(wrap_text=True, horizontal='center', vertical='center')
                    for litera in ["A", "B", "C", "D", "E"]:
                        new_sheet[f"{litera}3"].alignment = style_center
                        new_sheet[f"{litera}3"].font = Font(bold=True)
                        new_sheet[f"{litera}3"].border = thin_border_
                        new_sheet.column_dimensions[litera].width = 20
                    new_sheet.row_dimensions[3].height = 32

                    
                    table = []
                    for component in components:
                        line = []
                        line.append(component)
                        line.append(components[component][2])
                        if components[component][0] == 1:
                            line.append('Да')
                        elif components[component][0] == -1:
                            line.append('Нет')
                        else:
                            line.append('Не проверено')
                        if components[component][1] == 1:
                            line.append('Да')
                        elif components[component][1] == -1:
                            line.append('Нет')
                        else:
                            line.append('Не проверено')
                        if components[component][5] == 1:
                            line.append('Да')
                        else:
                            line.append('Нет')
                        table.append(line)
                    row_i = 4
                    for line in table:
                        for index, item in enumerate(line, start=1):
                            new_sheet.cell(row=row_i, column=index, value=item).alignment = style_center
                            new_sheet.cell(row=row_i, column=index).border = thin_border_
                            new_sheet.row_dimensions[row_i].height = 21
                            match item:
                                case "Не проверено":
                                    color_ = PatternFill(start_color=config['COLORS']['YELLOW'], end_color=config['COLORS']['YELLOW'], fill_type="solid")
                                case "Нет":
                                    color_ = PatternFill(start_color=config['COLORS']['RED'], end_color=config['COLORS']['RED'], fill_type="solid")
                                case "Да":
                                    color_ = PatternFill(start_color=config['COLORS']['GREEN'], end_color=config['COLORS']['GREEN'], fill_type="solid")
                            if item in ["Да", "Нет", "Не проверено"]:
                                new_sheet.cell(row=row_i, column=index).fill = color_

                        row_i += 1
                    # for component in components:
                    #     #print(component)
                    #     new_sheet.cell(row=row_i, column=1, value=component).alignment = style_center
                    #     new_sheet.cell(row=row_i, column=1).border = thin_border_

                    #     for j in range(len(components[component])):
                    #         new_sheet.cell(row=row_i, column=j+2, value=components[component][j]).alignment = style_center
                    #         new_sheet.cell(row=row_i, column=j+2).border = thin_border_
                    #         new_sheet.row_dimensions[row_i].height = 21
                    #     row_i += 1
                    add_outer_border(new_sheet, f'A3:E{row_i-1}')
                    set_column_autowidth(new_sheet, ["B"], reserve=1.2)

                    hyperlink_cell2 = ws.cell(row=row_id, column=col_idx)  
                    #hyperlink_cell2.value = str(eval(row_bom['SMT Левенштейна'])[0][1]).replace("/", "").replace(".", "_") #"Похожие"
                    hyperlink_cell2.hyperlink = f"#'{value}'!A1" 
                    hyperlink_cell2.font = Font(underline="single", color="0563C1") 
                
            row_id += 1

    add_outer_border(ws, f'A3:H{row_id-1}')
    try:
        wb.save(path_dashboard_file)
    except Exception as e:
        log(f"ERROR: {e}")


def get_dashboard_html(config):
    """
    Создает таблицу дашборда в формате html с подготовленной версткой для мессенджера.
    
    :param config: Конфигурация
    """

    data_all = load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    data = data_all["PROJECTS"]
    

    if len(data) > 1:

        data_dataframe = get_dataframe(data)
        #df1 = data_dataframe.iloc[:, 0:5]

        df2 = data_dataframe.iloc[:, [0, 1, 5, 6, 7]]

        table_html = "Таблица состояний<br><br>" + df2.to_html(index=False) + f"<br>Актуально на: {datetime.datetime.now().strftime('%Y.%m.%d %H:%M')}"

        table_html = table_html.replace('<table border="1" class="dataframe">', '<table class="dataframe" style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;"> <col span="5"> <col style="width: 100px;"> <col style="width: 0%;"> <col style="width: 100px;"> <col style="width: 0%;"> <col style="width: 100px;">')
        table_html = table_html.replace('<th>', '<th style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;">')
        table_html = table_html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
        table_html = table_html.replace('<tr>','<tr style="text-align: center;">')
        table_html = table_html.replace('<td>','<td style="border: 1px solid rgb(155, 155, 155); overflow-wrap: break-word; border-collapse: collapse;">')
    else:
        table_html = "Нет текущих проектов."

    return table_html