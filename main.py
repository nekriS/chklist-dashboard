from system import log, create_thread_task, options
from messanger.PostLink import PostLinkClient
from core.listener import getListener
from core.services import checkStatus, updateDashboard, get_tmp_parts_from_db, check_function
from web.server import serverStart
from web.utils import get_computer_link

if __name__ == "__main__":
    
    log("START")
    options = options()
    options.print_all_options()
    options.create_folders()

    client = PostLinkClient(options.config["BOT"]['API_BASE_URL'], options.config["BOT"]['WS_URL'], silent=False, listener=getListener(options.config), logger=log)
    client.download_folder = f"{options.config["GENERAL"]['DEFAULT_PATH']}{options.config["GENERAL"]['NAME_FOLDER_UPLOADS']}"
    try:
        client.ensure_connected()
    except Exception as e:
        log(f"ERROR: {e}")

    if client.is_auth:
        while not client.is_connected:
            pass
        else:
            client.subscribe("/topic/messages")


    thread, destroy_check_function = create_thread_task(options.config["GENERAL"]['CHECK_TIMEOUT'], check_function, options.config, client)
    thread_2, destroy_tmp_parts = create_thread_task(options.config["DB"]['CHECK_TIMEOUT'], get_tmp_parts_from_db, options.config)
    thread_3, destroy_web_server = serverStart(port=2000, debug=False, logger=log)

    while True:
        command = input("")
        log(command)
        req_massive = command.split(" ")
        match req_massive[0]:
            case "exit":
                destroy_check_function.set()
                destroy_tmp_parts.set()
                destroy_web_server()
                break
            case "reload":
                destroy_check_function.set()
                destroy_tmp_parts.set()
                destroy_web_server()
                options.update()
                options.print_all_options()
                options.create_folders()
                thread, destroy_check_function = create_thread_task(options.config["GENERAL"]['CHECK_TIMEOUT'], check_function, options.config, client)
                thread_2, destroy_tmp_parts = create_thread_task(options.config["DB"]['CHECK_TIMEOUT'], get_tmp_parts_from_db, options.config)
                thread_3, destroy_web_server = serverStart(port=2000, debug=False, logger=log)
            case "update":
                checkStatus(client, options.config)
                updateDashboard(options.config)
            case "connect":
                get_tmp_parts_from_db(options.config)
            case "send_message":
                client.send_message(target_id=int(req_massive[1]), target_type="person", text=" ".join(req_massive[2:]))
            case "try":
                try:
                    client.ensure_connected()
                    while not client.is_connected:
                        pass
                    else:
                        client.subscribe("/topic/messages")
                except:
                    log("ERROR CONNECT")
            case "web":
                print(get_computer_link())