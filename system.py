from config import NAME_CONFIG_FILE, DEFAULT_CONFIG
import configparser
import datetime
import os
import threading
import json

class options():
    def __init__(self,
                 config = configparser.ConfigParser()):
        self.config = config

        if os.path.exists(NAME_CONFIG_FILE):
            self.update()
        else:
            for section, value in DEFAULT_CONFIG.items():
                self.config[section] = value

            with open(NAME_CONFIG_FILE, "w", encoding="utf-8") as f:
                self.config.write(f)

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


def log(text):
    if text != "" and text != " ":
        # Получаем сегодняшнюю дату в формате YYYY-MM-DD
        today_date = datetime.datetime.now().strftime('%Y-%m-%d')

        # Создаем путь к каталогу и файлу
        log_directory = 'LOG'
        log_file_name = f'log_{today_date}.txt'
        log_file_path = os.path.join(log_directory, log_file_name)

        # Проверяем наличие каталога "log" и создаем его, если его нет
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)

        # Формируем строку для записи: "дата-время > текст"
        current_time = datetime.datetime.now().strftime('%H:%M:%S')
        log_entry = f"{today_date} {current_time} > {text}"

        # Проверяем existence файла и записываем данные
        with open(log_file_path, 'a', encoding='utf-8') as file:
            file.write(log_entry+"\n")

        try:
            print(log_entry)
        except:
            print(log_entry.encode("utf-8"))


def create_thread_task(period, function, *args, **kwargs):
    """
    Выполняет функцию func каждые period секунд в отдельном потоке.
    
    :param func: функция, которую нужно выполнять
    :param args: позиционные аргументы для функции
    :param kwargs: именованные аргументы для функции
    """
    def wrapper(destroy_event):
        while not destroy_event.is_set():
            try:
                function(*args, **kwargs)
            except Exception as e:
                log(f"ERROR with function: {function} {args} {kwargs} {e}")
            destroy_event.wait(int(period))

    destroy_event = threading.Event()
    thread = threading.Thread(target=wrapper, args=(destroy_event,), daemon=False)
    thread.start()
    return thread, destroy_event


def load_data(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except:
        log(f'File {path} was created!')
        data = {}
        notifications = {}
        projects = {}
        projects["_COUNTER"] = 0
        data["PROJECTS"] = projects
        data["NOTIFICATIONS"] = notifications
        with open(path, "w", encoding='utf-8') as f:
            f.write(json.dumps(data, indent=4))
    finally:
        f.close()
    return data

def save_data(data, path):
    with open(path, "w", encoding='utf-8') as f:
        f.write(json.dumps(data, indent=4))
    f.close()