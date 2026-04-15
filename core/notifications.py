import datetime
from system import log
from .utils import check_date_diff

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
            return data_all
        for target in targets:
            if f"{project}-{target}-{type_n}" not in list(data_all["NOTIFICATIONS"].keys()):
                notification = {}
                notification["PROJECT_NAME"] = project
                notification["TARGET"] = target
                notification["STEP"] = step
                notification["NEXT_TIME"] = f"{(datetime.datetime.now() + datetime.timedelta(days=notification["STEP"])).strftime('%Y.%m.%d')}"
                notification["COUNT"] = count
                notification["TYPE"] = type_n
                data_all["NOTIFICATIONS"][f"{project}-{target}-{type_n}"] = notification
                log(f"Notification {project}-{target}-{type_n} was created!")
        return data_all
    except Exception as e:
        log(f"create_task_notification {e}")

def delete_task_notification(config, data, project, target, type_n):

    if isinstance(target, int):
        targets = [target]
    elif isinstance(target, str):
        if target in config["USER_RIGHTS"].keys():
            targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
        elif target in config["USER_ALIASE"].keys():
            targets = [int(config["USER_ALIASE"][target])]
        #targets = []
    else:
        return data
    
    for target in targets:
        if f"{project}-{target}-{type_n}" in list(data["NOTIFICATIONS"].keys()):
            del data["NOTIFICATIONS"][f"{project}-{target}-{type_n}"]
            log(f"Notification {project}-{target}-{type_n} was deleted!")

    return data

def makeBorder(text):
    return  f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;">{text}</div>'

def send_notifications(client, notifications):

    for notice in notifications.keys():

        if abs(check_date_diff(datetime.datetime.strptime(notifications[notice]["NEXT_TIME"], "%Y.%m.%d"), datetime.datetime.now())) < 1:

            notifications[notice]
            match notifications[notice]["TYPE"]:
                case 0:
                    text = makeBorder(f"<strong>Внимание!</strong><br><br>Для чеклиста проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong> необходимо распределить проверяющих!")
                case 1:
                    text = makeBorder(f"<strong>Внимание!</strong><br><br>Проверьте, пожалуйста, схем символы проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong>.")   
                case 2:
                    text = makeBorder(f"<strong>Внимание!</strong><br><br>Проверьте, пожалуйста, посадочные проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong>.")
                case 3:
                    text = makeBorder(f"<strong>Внимание!</strong><br><br>Внесит, пожалуйста, правки в проект <strong>{notifications[notice]["PROJECT_NAME"]}</strong>.")
                case 4:
                    text = makeBorder(f"<strong>Внимание!</strong><br><br>В проекте <strong>{notifications[notice]["PROJECT_NAME"]}</strong> все компоненты готовы к переводу в базу!")
                case _:
                    log(f"Notification is bad. Can't resolve type: {notifications[notice]["TYPE"]}, notice: {notice}")
                    text = None

                # case 0:
                #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Для чеклиста проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong> необходимо распределить проверяющих!</div>'
                # case 1:
                #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Проверьте, пожалуйста, схем символы проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong>.</div>'
                # case 2:
                #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Проверьте, пожалуйста, посадочные проекта <strong>{notifications[notice]["PROJECT_NAME"]}</strong>.</div>'
                # case 4:
                #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>В проекте <strong>{notifications[notice]["PROJECT_NAME"]}</strong> все компоненты готовы к переводу в базу!</div>'
                #     #text = f"<strong>Внимание!</strong><br><br>В проекте <mark>{notifications[notice]["PROJECT_NAME"]}</mark> все компоненты готовы к переводу в базу!"
                #     #f"<strong>Внимание!</strong><br><br>В проекте <mark>CUBESAT_TRANSCEIVER_2_25033_R1</mark> все компоненты готовы к переводу в базу!"
                # case _:
                #     text = None
                
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