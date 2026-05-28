from system import log, loadData, saveData
from .notifications import send_notifications
from .combine import combine_checklists
#from .states import checkStatus
from worksheets.services import drawXDashboard
from web.services import drawHDashboard
import pandas as pd
import datetime
from pyodbc import connect
from pathlib import Path
import shutil
import os



from .notifications import create_task_notification, delete_task_notification
from .utils import check_date_diff


def getStateByNumber(stateNumber: int) -> str:
    match stateNumber:
        case 0 | 1 | 2:
            return "Новый"
        case 3 | 4 | 401 | 5:
            return "Идет проверка"
        case 6:
            return "К переводу"
        case 8:
            return "Готов вручную"
        case 9:
            return "Готов номинально"
        case 10:
            return "Готов"
        case _:
            return ""

def sendLastFile(client, config, target, text, file):
    try:
        if isinstance(target, int):
            targets = [target]
        elif isinstance(target, str):
            if target in config["USER_RIGHTS"].keys():
                targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
            elif target in config["USER_ALIASE"].keys():
                targets = [int(config["USER_ALIASE"][target])]
            #targets = []
        else:
            targets = []
        
        for id in targets:
            client.send_file(target_id = id, filepath = file, target_type="person", message=text)

    except Exception as e:
        log(f"Can not send last file: target={target}, error={e}")

def getDataframe(projects: dict):
    dataList = []
    for project in projects.keys():
        if not(project == "_COUNTER"):
            if not projects[project]["VISIBLE"]:
                continue
            
            row_data = []
            row_data.append(projects[project]["NUMBER"])
            row_data.append(projects[project]["PROJECT_NAME"])
            row_data.append(projects[project]["DEVELOPER"])
            row_data.append(getStateByNumber(int(projects[project]["STATE"])))
            row_data.append(datetime.datetime.strptime(projects[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"))
            row_data.append(f"{projects[project]["BD"]:>3} / {projects[project]["COUNT"]:<3}")
            row_data.append(f"{projects[project]["FP"]:>3} / {projects[project]["COUNT"]:<3}")
            row_data.append(f"{projects[project]["noTMP"]:>3} / {projects[project]["COUNT"]:<3}")
            dataList.append(row_data)

    data_df = pd.DataFrame(dataList)
    status_order = ["Новый", "Идет проверка", "К переводу", "Готов", "Готов номинально", "Готов вручную"][::-1]
    data_df[3] = pd.Categorical(data_df[3], categories=status_order, ordered=True)
    projectsDataframe = data_df.sort_values(3).reset_index(drop=True)
    projectsDataframe.columns = ['ID', 'Имя проекта', 'Разработчик', 'Статус', 'Последнее обновление', 'БД+Символ', 'Посадочное', 'Переведен в постоянные']

    return projectsDataframe

def getSubtables(config, projects: dict):
    subtables = {}
    for project in projects.keys():
        if not(project == "_COUNTER"):
            subtable = getSubTable(config, projects, project)
            subtables[project] = subtable
    return subtables

def getSubTable(config, projects: dict, project: str) -> list:
    components = getComponents(config, projects, project)
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

    return table

def getComponents(config, data, project):
    components = {}
    try:
        TMP_PARTS = pd.read_csv(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_BD']}").dropna(subset=["NAME"])['TMP_NAME'].values
    except:
        TMP_PARTS = []

    for path in [data[project]["PATH"][-1]]:

        with pd.option_context('future.no_silent_downcasting', True):
            df = pd.read_excel(f"{path}", header=None, names=["B", "C", "D", "E", "F", "J", "K"], usecols="B, C, D, E, F, J, K",skiprows=7).replace(to_replace={'Не проверено': 0, 'Да': 1, 'Нет': -1})
            column_order = ["B", "C", "D", "E", "F", "J", "K"]
            df = df[column_order].fillna('_NONE')
            df = df.replace(' ', '_NONE')
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
                if lines[5] != 0 and lines[5] != "_NONE":
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

def get_list_checkers(components):

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
                pass
            #del sch_checkers[component[3]]
        else:
            return False
        
        if component[4] != "_NONE":

            if component[4] not in pcb_checkers.keys():
                pcb_checkers[component[4]] = []

            if component[1] == 0 and component[5] != 1:
                pcb_checkers[component[4]].append(key)
            else:
                pass

            #    del pcb_checkers[component[4]]
        else:
            return False
    
    checkers = {}
    checkers["sch"] = sch_checkers
    checkers["pcb"] = pcb_checkers

    return checkers


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
        #data = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
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

        #saveData(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
        log(f"Project {project_name} was added!")
    except Exception as e:
        log(f"ERROR: {e}")
    #update_dashboard(config)


def get_dashboard_html(config):
    """
    Создает таблицу дашборда в формате html с подготовленной версткой для мессенджера.
    
    :param config: Конфигурация
    """

    data_all = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    data = data_all["PROJECTS"]
    

    if len(data) > 1:

        data_dataframe = getDataframe(data)
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

def updateDashboard(config):

    data = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    projects = data["PROJECTS"]

    dataframe = getDataframe(projects)
    subtables = getSubtables(config, projects)

    drawXDashboard(config, dataframe, subtables)
    drawHDashboard(config, dataframe, subtables)


def check_function(config, client):

    data = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    projects = data["PROJECTS"]
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

                if not(project_name in projects):
                    add_new_checklist(projects, new_file_path)
                    
                else:
                    last_file = projects[project_name]["PATH"][-1]
                    projects[project_name]["PATH"].append(new_file_path)
                    projects[project_name]["COUNT"] = max(projects[project_name]["COUNT"], len(pd.read_excel(f"{new_file_path}", header=None, usecols="A",skiprows=7).fillna("_NONE").values))

                    corr_file_path = f"{path_work_folder}/{project_name}_{today_date}_{current_time}_corr.xlsx".replace(" ", "_")
                    combine_checklists(last_file, new_file_path, corr_file_path)
                    projects[project_name]["PATH"].append(corr_file_path)
                    #client.send_message()
                    
                    log(f"New PATH in {project_name} was added!")
            except:
                pass
        
        #update_dashboard(config)
    
    # for project in data:
    #     if not(project == "_COUNTER"):
    #         if "Готов" in data[project]["STATUS"]:
    #             if check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_INVISIBLE"]) and bool(data[project]["VISIBLE"]) == True:
    #                 data[project]["VISIBLE"] = False
    #                 log(f"Project {project} has become invisible!")
    #             if check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_DELETE"]) and bool(data[project]["VISIBLE"]) == False:
    #                 delete_project(project)

    #         if data[project]["noTMP"] < data[project]["COUNT"] and data[project]["BD"] == data[project]["COUNT"] and data[project]["FP"] == data[project]["COUNT"]:
    #             data_all = create_task_notification(config, data_all, project, "lib_manager", 0)
    #         elif (data[project]["noTMP"] == data[project]["COUNT"] and data[project]["BD"] == data[project]["COUNT"] and data[project]["FP"] == data[project]["COUNT"]) or data[project]["noTMP"] == 0:
    #             data_all = delete_task_notification(config, data_all, project, "lib_manager", 0)

    
    data["NOTIFICATIONS"] = send_notifications(client, data["NOTIFICATIONS"])

    data["PROJECTS"] = projects
    saveData(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    try:
        checkStatus(client, config)
    except Exception as e:
        log(f"[WARN] Failed to check status: {e}")
    try:
        updateDashboard(config)
    except Exception as e:
        log(f"[WARN] Failed to update dashboard: {e}")








def delete_project(project):
    pass

def checkStatus(client, config):

    data = loadData(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    projects = data["PROJECTS"]

    for project in projects.keys():
        if not(project == "_COUNTER"):



            state = projects[project]["STATE"]
            path_last_file = projects[project]["PATH"][-1]

            developer = projects[project]["DEVELOPER"]

            components = getComponents(config, projects, project)

            yes_bd_schem = 0
            no_bd_schem = 0
            pcb_yes = 0
            pcb_no = 0
            noTMP = 0
            for comp in components:
                if components[comp][0] == 1:
                    yes_bd_schem += 1
                elif components[comp][0] == -1:
                    no_bd_schem += 1
                if components[comp][1] == 1:
                    pcb_yes += 1
                elif components[comp][1] == -1:
                    pcb_no += 1
                noTMP += (components[comp][5])

            if noTMP == len(components) and (yes_bd_schem != len(components) or pcb_yes != len(components)):
                projects[project]["STATE"] = 9

            if state >= 9 and state < 11:
                delete_task_notification(config, data, project, "all", 0)
                delete_task_notification(config, data, project, "all", 1)
                delete_task_notification(config, data, project, "all", 2)
                delete_task_notification(config, data, project, "all", 3)
                delete_task_notification(config, data, project, "all", 4)

            match state:
                case 0: # Новый проект

                    message = f"""
Добавлен новый проект для проверки! Распределите их, пожалуйста, между проверяющими.<br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество компонентов: {len(components)}<br>
<br>
После заполнения полей с проверяющими, пришлите файл в этот чат. Рассылать проверяющим самостоятельно не требуется.
"""
                    sendLastFile(client, config, "lib_manager", message, path_last_file)

                    create_task_notification(config, data, project, "lib_manager", 0, step=7)

                    projects[project]["STATE"] = 1


                case 101: # Добавлен новый компонент в текущий проект
                    message = f"""
В действующем проекте были добавлены новые компоненты! Распределите их, пожалуйста, между проверяющими.<br>

<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество компонентов: {len(components)}<br>
<br>
После заполнения полей с проверяющими, пришлите файл в этот чат. Рассылать проверяющим самостоятельно не требуется.
"""
                    sendLastFile(client, config, "lib_manager", message, path_last_file)
                    create_task_notification(config, data, project, "lib_manager", 0, step=7)
                    projects[project]["STATE"] = 1
                
                case 1: # Компонент отослан библиотекарю на распределение

                    checkers = get_list_checkers(components)
                    if checkers:
                        delete_task_notification(config, data, project, "lib_manager", 0)
                        projects[project]["STATE"] = 2

                case 2: # Получен файл от библиотекаря

                    checkers = get_list_checkers(components)
                    if checkers:
                        for checker in checkers["sch"].keys():
                            if len(checkers["sch"][checker]) > 0:

                                message = f"""
Вы были назначены проверяющим схем символов компонентов!<br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["sch"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                sendLastFile(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 1, step=7)

                        for checker in checkers["pcb"].keys():
                            if len(checkers["pcb"][checker]) > 0:

                                message = f"""
Вы были назначены проверяющим посадочных компонентов!<br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["pcb"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                sendLastFile(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 2, step=7)
                        
                        projects[project]["STATE"] = 3

                    else:
                        projects[project]["STATE"] = 101

                case 3: # Файлы разосланы исполнителям, но никто не отослал файл

                    checkers = get_list_checkers(components)
                    if checkers:

                        for checker in checkers["sch"].keys():
                            if len(checkers["sch"][checker]) == 0:
                                projects[project]["STATE"] = 4

                        for checker in checkers["pcb"].keys():
                            if len(checkers["pcb"][checker]) == 0:
                                projects[project]["STATE"] = 4


                case 4: # Идет проверка, кто-то прислал файл
                    flag_all_check = True
                    checkers = get_list_checkers(components)
                    if checkers:

                        for checker in checkers["sch"].keys():
                            if len(checkers["sch"][checker]) == 0:
                                delete_task_notification(config, data, project, checker, 1)
                            else:
                                flag_all_check = False

                        for checker in checkers["pcb"].keys():
                            if len(checkers["pcb"][checker]) == 0:
                                delete_task_notification(config, data, project, checker, 2) 
                            else:
                                flag_all_check = False

                    else:
                        projects[project]["STATE"] = 101
                    print(f"flag_all_check {flag_all_check}")
                    if flag_all_check and yes_bd_schem == len(components) and pcb_yes == len(components):
                        projects[project]["STATE"] = 5
                        delete_task_notification(config, data, project, developer, 3)
                    elif flag_all_check and (yes_bd_schem != len(components) or pcb_yes != len(components)):
                        
                        message = f"""
Проект проверен, есть замечания. Исправьте, пожалуйста.<br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
<br>
После окончания правок, пожалуйста, пришлите свеже сгенерированный файл. Если Вы не согласны с некоторыми замечаниями, обсудите их с проверяющим. Если нужно исправить "Нет" на "Да", то сначала отправьте полученный файл с исправлением, а затем свеже сгенерированный чеклист.
"""
                        sendLastFile(client, config, developer, message, path_last_file)
                        create_task_notification(config, data, project, developer, 3)
                        projects[project]["STATE"] = 401

                    # elif not(flag_all_check):
                    #     delete_task_notification(config, data, project, developer, 3)
                    # if yes_bd_schem == len(components) and pcb_yes == len(components):
                    #     projects[project]["STATE"] = 5
                    #     delete_task_notification(config, data, project, developer, 3)
                    # else:
                    #     create_task_notification(config, data, project, developer, 3)

                case 401: #доработка

                    checkers = get_list_checkers(components)

                    a = 0
                    b = 0

                    if checkers and no_bd_schem == 0 and pcb_no == 0:

                        delete_task_notification(config, data, project, developer, 3)

                        for checker in checkers["sch"].keys():
                            b += 1
                            if len(checkers["sch"][checker]) == 0:
                                delete_task_notification(config, data, project, checker, 1)
                                a += 1
                            else:
                                message = f"""
В одном из проектов были внесены правки, проверьте, пожалуйста! <br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["sch"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                sendLastFile(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 1)
                                projects[project]["STATE"] = 4

                        for checker in checkers["pcb"].keys():
                            b += 1
                            if len(checkers["pcb"][checker]) == 0:
                                delete_task_notification(config, data, project, checker, 2)
                                a += 1
                            else:
                                message = f"""
В одном из проектов были внесены правки, проверьте, пожалуйста! <br>
<br><br>
Название проекта: {project}<br>
Разработчик: {developer}<br>
Количество ваших компонентов: {len(checkers["pcb"][checker])}<br>
<br>
После окончания проверки, пожалуйста, пришлите проверенный файл в данный чат. В файле не должно остаться компонентов со статусом "Не проверено".
"""
                                sendLastFile(client, config, checker, message, path_last_file)
                                create_task_notification(config, data, project, checker, 2)
                                projects[project]["STATE"] = 4
                        
                        if a == b:
                            projects[project]["STATE"] = 5
                                        

                case 5: # Все проверено
                    checkers = get_list_checkers(components)
                    if checkers:
                        
                        message = f"""
В проекте {project} все компоненты готовы к добавлению в базу! <br>
<br><br>
Разработчик: {developer}<br>
Количество компонентов: {len(components)}<br>
<br>
Дополнительных действий не требуется.
"""
                        sendLastFile(client, config, "lib_manager", message, path_last_file)
                        create_task_notification(config, data, project, "lib_manager", 4, step=7)
                        projects[project]["STATE"] = 6

                    else:
                        projects[project]["STATE"] = 101

                case 6:
                    if noTMP == len(components):

                        message = f"""
Все компоненты в проекте {project} добавлены в базу!"""
                        
                        delete_task_notification(config, data, project, "lib_manager", 4)
                        sendLastFile(client, config, developer, message, path_last_file)
                        #delete_task_notification(config, data, project, "lib_manager", 3) 
                        projects[project]["STATE"] = 10

                # case 8: # Готово вручную
                #     pass 

                # case 9: # Готово номинально, если все компоненты уже в базе
                #     pass
                
                case 8 | 9 | 10: # Готово

                    if check_date_diff(datetime.datetime.strptime(projects[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_INVISIBLE"]) and bool(projects[project]["VISIBLE"]) == True:
                        projects[project]["VISIBLE"] = False
                        log(f"Project {project} has become invisible!")
                    if check_date_diff(datetime.datetime.strptime(projects[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now()) > int(config["GENERAL"]["PERIOD_TO_DELETE"]) and bool(projects[project]["VISIBLE"]) == False:
                        delete_project(project)
            
            if projects[project]["BD"] != yes_bd_schem:
                projects[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            projects[project]["BD"] = yes_bd_schem

            if projects[project]["FP"] != pcb_yes:
                projects[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            projects[project]["FP"] = pcb_yes

            if projects[project]["COUNT"] != len(components):
                projects[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            projects[project]["COUNT"] = len(components)

            if projects[project]["noTMP"] != noTMP:
                projects[project]["LASTUPDATE"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            projects[project]["noTMP"] = noTMP

    data["PROJECTS"] = projects
    saveData(data, f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")