from system import log
from .utils import get_date_difference
from datetime import datetime as time

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

def send_message_with_last_file(messanger, config, target, message, file):
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
            messanger.send_file(target_id = id, 
                                filepath = file, 
                                target_type = "person", 
                                message = message)

    except Exception as e:
        log(f"Can not send last file: target={target}, error={e}")

def get_targets(config, target):
    if isinstance(target, int):
        return [target]
    elif isinstance(target, str):
        if target == "all":
            targets = []
            for tar in config["USER_ALIASE"].keys():
                targets.append(int(config["USER_ALIASE"][tar]))
            return targets
        elif target in config["USER_RIGHTS"].keys():
            return [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]

        elif target in config["USER_ALIASE"].keys():
            return [int(config["USER_ALIASE"][target])]
        else:
            return []
    else:
        return []

def check_state(dashboard, project, components):

    state = project["state"]
    path_last_file = project["path"].split("|")[-1]
    developer = project["developer"]
    project_name = project["project_name"]

    no_bd_schem = project["quantity"] - project["bd"]
    no_pcb = project["quantity"] - project["fp"]

    if state >= 8 and state < 11:
        for notification_type in [0, 1, 2, 3, 4]:
            dashboard.notifications.delete_notification(project_name, get_targets(dashboard.config, "all"), notification_type)

    match state:
        case 0: # Новый проект
            message = dashboard.messages["NEW_PROJECT"].format(project = project_name, 
                                                               developer = developer, 
                                                               count_components = len(components))

            send_message_with_last_file(dashboard.messanger, dashboard.config, "lib_manager", message, path_last_file)
            dashboard.notifications.create_notification(project_name, 
                                                        targets = get_targets(dashboard.config, "lib_manager"), 
                                                        notification_type = 0, 
                                                        step = 7)
            project["state"] = 1


        case 101: # Добавлен новый компонент в текущий проект
            message = dashboard.messages["NEW_COMPONENT"].format(project = project_name, 
                                                                 developer = developer, 
                                                                 count_components = len(components))

            send_message_with_last_file(dashboard.messanger, dashboard.config, "lib_manager", message, path_last_file)
            dashboard.notifications.create_notification(project_name, 
                                                        targets = get_targets(dashboard.config, "lib_manager"), 
                                                        notification_type = 0, 
                                                        step = 7)
            project["state"] = 1
        
        case 1: # Компонент отослан библиотекарю на распределение
            checkers = get_list_checkers(components)
            if checkers:
                dashboard.notifications.delete_notification(project_name, 
                                                            targets = get_targets(dashboard.config, "lib_manager"), 
                                                            notification_type = 0)
                project["state"] = 2

        case 2: # Получен файл от библиотекаря

            checkers = get_list_checkers(components)
            if checkers:
                for checker in checkers["sch"].keys():
                    if len(checkers["sch"][checker]) > 0:
                        message = dashboard.messages["CHECK_SCHEM"].format(project = project_name, 
                                                                           developer = developer, 
                                                                           count_components = len(checkers["sch"][checker]))
                        send_message_with_last_file(dashboard.messanger, dashboard.config, checker, message, path_last_file)
                        dashboard.notifications.create_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 1, 
                                                                    step = 7)

                for checker in checkers["pcb"].keys():
                    if len(checkers["pcb"][checker]) > 0:
                        message = dashboard.messages["CHECK_PCB"].format(project = project_name, 
                                                                         developer = developer, 
                                                                         count_components = len(checkers["pcb"][checker]))
                        send_message_with_last_file(dashboard.messanger, dashboard.config, checker, message, path_last_file)
                        dashboard.notifications.create_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 2, 
                                                                    step = 7)
                project["state"] = 3
            else:
                project["state"] = 101

        case 3: # Файлы разосланы исполнителям, но никто не отослал файл

            checkers = get_list_checkers(components)
            if checkers:

                for checker in checkers["sch"].keys():
                    if len(checkers["sch"][checker]) == 0:
                        project["state"] = 4

                for checker in checkers["pcb"].keys():
                    if len(checkers["pcb"][checker]) == 0:
                        project["state"] = 4


        case 4: # Идет проверка, кто-то прислал файл
            flag_all_check = True
            checkers = get_list_checkers(components)
            if checkers:

                for checker in checkers["sch"].keys():
                    if len(checkers["sch"][checker]) == 0:
                        dashboard.notifications.delete_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 1)
                    else:
                        flag_all_check = False

                for checker in checkers["pcb"].keys():
                    if len(checkers["pcb"][checker]) == 0:
                        dashboard.notifications.delete_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 2)
                    else:
                        flag_all_check = False

            else:
                project["state"] = 101

            if flag_all_check and project["bd"] == len(components) and project["fb"] == len(components):
                project["state"] = 5
                dashboard.notifications.delete_notification(project_name, 
                                                            targets = get_targets(dashboard.config, developer), 
                                                            notification_type = 3)

            elif flag_all_check and (project["bd"] != len(components) or project["fb"] != len(components)):
                message = dashboard.messages["MISTAKES"].format(project, developer)
                
                send_message_with_last_file(dashboard.messanger, dashboard.config, developer, message, path_last_file)
                dashboard.notifications.create_notification(project_name, 
                                                            targets = get_targets(dashboard.config, developer), 
                                                            notification_type = 3, 
                                                            step = 7)
                project["state"] = 401

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

            if checkers and no_bd_schem == 0 and no_pcb == 0:

                dashboard.notifications.delete_notification(project_name, 
                                                            targets = get_targets(dashboard.config, developer), 
                                                            notification_type = 3)

                for checker in checkers["sch"].keys():
                    b += 1
                    if len(checkers["sch"][checker]) == 0:
                        dashboard.notifications.delete_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, developer), 
                                                                    notification_type = 3)
                        a += 1
                    else:
                        message = dashboard.messages["CHECK_MISTAKES"].format(project = project_name, 
                                                                              developer = developer, 
                                                                              count_components = len(checkers["sch"][checker]))
                        send_message_with_last_file(dashboard.messanger, dashboard.config, checker, message, path_last_file)
                        dashboard.notifications.create_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 1, 
                                                                    step = 7)
                        project["state"] = 4

                for checker in checkers["pcb"].keys():
                    b += 1
                    if len(checkers["pcb"][checker]) == 0:
                        dashboard.notifications.delete_notification(project_name, 
                                                                    targets = developer, 
                                                                    notification_type = 2)
                        a += 1
                    else:
                        message = dashboard.messages["CHECK_MISTAKES"].format(project = project_name, 
                                                                              developer = developer, 
                                                                              count_components = len(checkers["pcb"][checker]))
                        send_message_with_last_file(dashboard.messanger, dashboard.config, checker, message, path_last_file)
                        dashboard.notifications.create_notification(project_name, 
                                                                    targets = get_targets(dashboard.config, checker), 
                                                                    notification_type = 2, 
                                                                    step = 7)
                        project["state"] = 4
                
                if a == b:
                    project["state"] = 5
                                

        case 5: # Все проверено
            checkers = get_list_checkers(components)
            if checkers:
                message = dashboard.messages["CHECKED"].format(project = project_name, 
                                                               developer = developer, 
                                                               count_components = len(components))
                send_message_with_last_file(dashboard.messanger, dashboard.config, "lib_manager", message, path_last_file)
                dashboard.notifications.create_notification(project_name, 
                                                            targets = get_targets(dashboard.config, "lib_manager"), 
                                                            notification_type = 4, 
                                                            step = 7)
                project["state"] = 6

            else:
                project["state"] = 101

        case 6:
            if project["noTMP"] == len(components):
                message = dashboard.messages["CHECKED"].format(project = project_name)
                dashboard.notifications.delete_notification(project_name, 
                                                            targets = get_targets(dashboard.config, "lib_manager"), 
                                                            notification_type = 4)
                send_message_with_last_file(dashboard.messanger, dashboard.config, developer, message, path_last_file)
                #delete_task_notification(config, data, project, "lib_manager", 3) 
                project["state"] = 10

        # case 8: # Готово вручную
        #     pass 

        # case 9: # Готово номинально, если все компоненты уже в базе
        #     pass
        
        case 8 | 9 | 10: # Готово

            if get_date_difference(time.strptime(project["updated_at"], "%Y-%m-%d %H:%M:%S"), time.now()) > int(dashboard.config["GENERAL"]["PERIOD_TO_INVISIBLE"]) and bool(project["visible"]) == True:
                project["visible"] = False
                log(f"Project {project_name} has become invisible!")
            # if check_date_diff(time.strptime(project["updated_at"], "%Y-%m-%d %H:%M:%S"), time.now()) > int(dashboard.config["GENERAL"]["PERIOD_TO_DELETE"]) and bool(project["visible"]) == False:
            #     delete_project(project)