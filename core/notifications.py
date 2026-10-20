# import datetime
# from system import log
# from .utils import check_date_diff

import sqlite3
from datetime import datetime as time
from datetime import timedelta as delta

BORDER = '<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;">{}</div>'

def check_date_diff(date1, date2):
    diff = (date2 - date1).days
    return diff



class notificationsObject:

    def __init__(self, db_path="notifications.db", log=None):
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.cursor = self.conn.cursor()
        self.logger = log
        self.prefix = "[N]"
        self.table_name = "notifications"
        self._initialization()

    def _log(self, message):
        if self.logger == None:
            print(f"{self.prefix} [{time.now()}] > {message}")
        else:
            self.logger(f"{self.prefix} {message}")

    def _initialization(self):
        self.cursor.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {self.table_name} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                notification_id TEXT,
                project_name TEXT,
                target_id INTEGER,
                step INTEGER,
                next_time TEXT,
                quantity INTEGER,
                type INTEGER,
                is_read INTEGER DEFAULT 0,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(notification_id, project_name)
            )
        """
        )
        self.conn.commit()

    def process_batch(self, batch: list[dict]):
        try:
            for item in batch:
                action = item.get("action", "save")

                if action == "save":
                    self.cursor.execute(f"""
                        INSERT INTO {self.table_name} (
                            notification_id, project_name, target_id, step, next_time, quantity, type, is_read, updated_at
                        ) VALUES (
                            ?, ?, ?, ?, ?, ?, ?, 0, STRFTIME('%Y-%m-%d %H:%M:%S', 'now', 'localtime')
                        )
                        ON CONFLICT(notification_id, project_name) DO UPDATE SET
                            target_id = excluded.target_id,
                            step = excluded.step,
                            next_time = excluded.next_time,
                            quantity = excluded.quantity,
                            type = excluded.type,
                            is_read = {self.table_name}.is_read,
                            updated_at = STRFTIME('%Y-%m-%d %H:%M:%S', 'now', 'localtime')
                    """, (
                        item["notification_id"],
                        item["project_name"],
                        item["target_id"],
                        item["step"],
                        item["next_time"],
                        item["quantity"],
                        item["type"]
                    ))

                elif action == "delete":
                    self.cursor.execute(f"""
                        DELETE FROM {self.table_name} 
                        WHERE notification_id = ? AND project_name = ?
                    """, (item["notification_id"], item["project_name"]))

            self.conn.commit()
            self._log(f"A batch of {len(batch)} items has been successfully processed.")

        except sqlite3.Error as e:
            self.conn.rollback()
            self._log(f"Error during transaction execution: {e}")
            raise e

    def has_notification(self, notification_id: str, project_name: str) -> bool:
        """
        Проверяет, существует ли уведомление с заданным notification_id в проекте.
        Возвращает True, если запись найдена, и False, если её нет.
        """
        try:
            # Используем EXISTS — это самый быстрый способ проверки в SQL,
            # так как база останавливает поиск сразу, как только найдет первое совпадение.
            self.cursor.execute(f"""
                SELECT EXISTS(
                    SELECT 1 FROM {self.table_name} 
                    WHERE notification_id = ? AND project_name = ?
                )
            """, (notification_id, project_name))
            
            # fetchone() вернет кортеж вида (1,) или (0,)
            result = self.cursor.fetchone()
            return bool(result[0])
            
        except sqlite3.Error as e:
            self._log(f"Error while checking notification existence: {e}")
            return False

    def get_all_active(self) -> list[dict]:
        """Возвращает все уведомления из таблицы в виде списка словарей."""
        try:
            self.cursor.execute(f"SELECT * FROM {self.table_name}")
            # Превращаем каждую строку sqlite3.Row в стандартный dict
            return [dict(row) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            self._log(f"Error fetching all active notifications: {e}")
            return []

    def close(self):
        self.conn.close()

    def create_notification(self, 
                            project_name: str, 
                            targets: list, 
                            notification_type: int, 
                            step: int = 3, 
                            quantity: int = 356):

        for target_id in targets:
            if not self.has_notification(f"{target_id}-{notification_type}", project_name):
                self.process_batch([{
                    "action":"save", 
                    "notification_id": f"{target_id}-{notification_type}",
                    "project_name": project_name,
                    "target_id": target_id,
                    "step": step,
                    "next_time": f"{(time.now() + delta(days=step)).strftime('%Y.%m.%d')}",
                    "quantity": quantity,
                    "type": notification_type
                    }])
                self._log(f"Notification for {project_name}, {target_id}-{notification_type} was created!")

    def delete_notification(self, project_name, targets, notification_type):

        for target_id in targets:
            if self.has_notification(f"{target_id}-{notification_type}", project_name):

                self.process_batch([{
                    "action":"delete", 
                    "notification_id": f"{target_id}-{notification_type}",
                    "project_name": project_name
                    }])
            
                self._log(f"Notification for {project_name}, {target_id}-{notification_type} was deleted!")

    def send_notifications(self, client, messages):

        notifications = self.get_all_active()

        for notice in notifications:
            if abs(check_date_diff(time.strptime(notice["next_time"], "%Y.%m.%d"), time.now())) < 1:

                match notice["notification_type"]:
                    case 0:
                        message = BORDER.format(messages["N_ASSING_INSPECTORS"].format(notice["project_name"]))
                    case 1:
                        message = BORDER.format(messages["N_CHECK_SCHEM"].format(notice["project_name"]))
                    case 2:
                        message = BORDER.format(messages["N_CHECK_PCB"].format(notice["project_name"]))
                    case 3:
                        message = BORDER.format(messages["N_CHECK_MISTAKES"].format(notice["project_name"]))
                    case 4:
                        message = BORDER.format(messages["N_CHECKED"].format(notice["project_name"]))
                    case _:
                        self._log(f"Notification is bad. Can't resolve type: {notice["notification_type"]}, notice: {notice}")
                        text = None

                if text is not None:
                    client.send_message(target_id=int(notice["target_id"]), target_type="person", text=message)
                    notice["quantity"] -= 1
                    if notice["quantity"] > 0:
                        notice["next_time"] = (time.strptime(notice["next_time"], "%Y.%m.%d") + delta(days=notice["step"])).strftime('%Y.%m.%d')
                        self.process_batch([notice])
                    else:
                        self.delete_notification(notice["project_name"], [notice["target_id"]], notice["notification_type"])


if __name__ == "__main__":
    db = notificationsObject()

    db.create_notification("ARGO", [789125], 0)
    
    db.create_notification("ARGO211", [789125, 456779], 0)

    db.delete_notification("ARGO21", [456779], 0)

    #print(db.get_all_active())


    # db.process_batch([{
    #     "action":"save", 
    #     "notification_id": "464644-ARGO-1",
    #     "project_name": "ARGO",
    #     "target_id": 788844,
    #     "step": 7,
    #     "next_time": time.now().strftime("%Y.%m.%d"),
    #     "quantity": 365,
    #     "type": 2
    #     }])

    # table = db.get_all_active()

    # print(table[0][5])
    # print(time.strptime(table[0][5], "%Y.%m.%d"))


# if isinstance(target, int):
#             targets = [target]
#         elif isinstance(target, str):
#             if target == "all":
#                 targets = []
#                 for tar in config["USER_ALIASE"].keys():
#                     targets.append(int(config["USER_ALIASE"][tar]))

#             elif target in config["USER_RIGHTS"].keys():
#                 targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]

#             elif target in config["USER_ALIASE"].keys():
#                 targets = [int(config["USER_ALIASE"][target])]
#             #targets = []
#         else:
#             return data

# def get_array_target_id(config, target):

#     if isinstance(target, int):
#         return [target]
#     elif isinstance(target, str):
#         if target in config["USER_RIGHTS"].keys():
#             targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
#         elif target in config["USER_ALIASE"].keys():
#             targets = [int(config["USER_ALIASE"][target])]
#         return targets
#     else:
#         return []




        # if f"{project}-{target}-{type_n}" not in list(data_all["NOTIFICATIONS"].keys()):
        #     notification = {}
        #     notification["PROJECT_NAME"] = project
        #     notification["TARGET"] = target
        #     notification["STEP"] = step
        #     notification["NEXT_TIME"] = f"{(datetime.datetime.now() + datetime.timedelta(days=notification["STEP"])).strftime('%Y.%m.%d')}"
        #     notification["COUNT"] = count
        #     notification["TYPE"] = type_n
        #     data_all["NOTIFICATIONS"][f"{project}-{target}-{type_n}"] = notification
        #     log(f"Notification {project}-{target}-{type_n} was created!")
    







# conn = sqlite3.connect("notifications.db")
# cursor = conn.cursor()

# cursor.execute(
#         """
#         CREATE TABLE IF NOT EXISTS user_notifications (
#             id INTEGER PRIMARY KEY AUTOINCREMENT,
#             notification_id TEXT,
#             project_name TEXT,
#             target_id INTEGER,
#             step INTEGER,
#             next_time INTEGER,
#             quantity INTEGER,
#             type INTEGER,
#             is_read INTEGER DEFAULT 0
#         )
#     """
#     )
# conn.commit()

# def create_task_notification(config, data_all, project, target, type_n, step=3, count=356):
#     """
#     Docstring для create_task_notification
    
#     :param config: Конфигурация
#     :param data_all: Данные
#     :param project: Проект
#     :param target: Цель, конкретный пользователь или группа
#     :param type_n: Тип
#     :param step: Шаг в днях (по умолчанию 3)
#     :param count: Количество (по умолчанию 356)
#     """

#     #data["NOTIFICATION"] = []
#     try:
#         if isinstance(target, int):
#             targets = [target]
#         elif isinstance(target, str):
#             if target in config["USER_RIGHTS"].keys():
#                 targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]
#             elif target in config["USER_ALIASE"].keys():
#                 targets = [int(config["USER_ALIASE"][target])]
#             #targets = []
#         else:
#             return data_all
#         for target in targets:
#             if f"{project}-{target}-{type_n}" not in list(data_all["NOTIFICATIONS"].keys()):
#                 notification = {}
#                 notification["PROJECT_NAME"] = project
#                 notification["TARGET"] = target
#                 notification["STEP"] = step
#                 notification["NEXT_TIME"] = f"{(datetime.datetime.now() + datetime.timedelta(days=notification["STEP"])).strftime('%Y.%m.%d')}"
#                 notification["COUNT"] = count
#                 notification["TYPE"] = type_n
#                 data_all["NOTIFICATIONS"][f"{project}-{target}-{type_n}"] = notification
#                 log(f"Notification {project}-{target}-{type_n} was created!")
#         return data_all
#     except Exception as e:
#         log(f"create_task_notification {e}")

# def delete_task_notification(config, data, project, target, type_n):

#     if isinstance(target, int):
#         targets = [target]
#     elif isinstance(target, str):
#         if target == "all":
#             targets = []
#             for tar in config["USER_ALIASE"].keys():
#                 targets.append(int(config["USER_ALIASE"][tar]))

#         elif target in config["USER_RIGHTS"].keys():
#             targets = [int(config["USER_ALIASE"][x.upper().strip().replace("'","")]) for x in config["USER_RIGHTS"][target.upper()].strip('[]').split(',')]

#         elif target in config["USER_ALIASE"].keys():
#             targets = [int(config["USER_ALIASE"][target])]
#         #targets = []
#     else:
#         return data
    
#     for target in targets:
#         if f"{project}-{target}-{type_n}" in list(data["NOTIFICATIONS"].keys()):
#             del data["NOTIFICATIONS"][f"{project}-{target}-{type_n}"]
#             log(f"Notification {project}-{target}-{type_n} was deleted!")

#     return data

# def makeBorder(text):
#     return  f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;">{text}</div>'

# def send_notifications(client, notifications):

#     for notice in notifications.keys():

#         if abs(check_date_diff(datetime.datetime.strptime(notifications[notice]["NEXT_TIME"], "%Y.%m.%d"), datetime.datetime.now())) < 1:

#             notifications[notice]
#             match notifications[notice]["TYPE"]:
#                 case 0:
#                     text = makeBorder(f"<strong>Внимание!</strong><br><br>Для чеклиста проекта <strong>{notice["project_name"]}</strong> необходимо распределить проверяющих!")
#                 case 1:
#                     text = makeBorder(f"<strong>Внимание!</strong><br><br>Проверьте, пожалуйста, схем символы проекта <strong>{notice["project_name"]}</strong>.")   
#                 case 2:
#                     text = makeBorder(f"<strong>Внимание!</strong><br><br>Проверьте, пожалуйста, посадочные проекта <strong>{notice["project_name"]}</strong>.")
#                 case 3:
#                     text = makeBorder(f"<strong>Внимание!</strong><br><br>Внесите, пожалуйста, правки в проект <strong>{notice["project_name"]}</strong>.")
#                 case 4:
#                     text = makeBorder(f"<strong>Внимание!</strong><br><br>В проекте <strong>{notice["project_name"]}</strong> все компоненты готовы к переводу в базу!")
#                 case _:
#                     log(f"Notification is bad. Can't resolve type: {notifications[notice]["TYPE"]}, notice: {notice}")
#                     text = None

#                 # case 0:
#                 #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Для чеклиста проекта <strong>{notice["project_name"]}</strong> необходимо распределить проверяющих!</div>'
#                 # case 1:
#                 #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Проверьте, пожалуйста, схем символы проекта <strong>{notice["project_name"]}</strong>.</div>'
#                 # case 2:
#                 #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>Проверьте, пожалуйста, посадочные проекта <strong>{notice["project_name"]}</strong>.</div>'
#                 # case 4:
#                 #     text = f'<div style="width: 300px; margin-top: 10px; margin-bottom: 10px; border: 3px solid yellow; padding-left: 10px; padding-right: 10px; padding-top: 10px; padding-bottom: 10px;"><strong>Внимание!</strong><br><br>В проекте <strong>{notice["project_name"]}</strong> все компоненты готовы к переводу в базу!</div>'
#                 #     #text = f"<strong>Внимание!</strong><br><br>В проекте <mark>{notice["project_name"]}</mark> все компоненты готовы к переводу в базу!"
#                 #     #f"<strong>Внимание!</strong><br><br>В проекте <mark>CUBESAT_TRANSCEIVER_2_25033_R1</mark> все компоненты готовы к переводу в базу!"
#                 # case _:
#                 #     text = None
                
#             if text is not None:
#                 client.send_message(target_id=int(notifications[notice]["TARGET"]), target_type="person", text=text)
#                 notifications[notice]["COUNT"] -= 1
#                 if notifications[notice]["COUNT"] > 0:
#                     #notifications[notice]["NEXT_TIME"] = notifications[notice]["NEXT_TIME"]
#                     #notifications[notice]["NEXT_TIME"]
#                     notifications[notice]["NEXT_TIME"] = (datetime.datetime.strptime(notifications[notice]["NEXT_TIME"], "%Y.%m.%d") + datetime.timedelta(days=notifications[notice]["STEP"])).strftime('%Y.%m.%d')
#                 else:
#                     del notifications[notice]
#                     #check_date_diff(datetime.datetime.strptime(data[project]["LASTUPDATE"], "%Y-%m-%d %H:%M:%S"), datetime.datetime.now())

#     return notifications