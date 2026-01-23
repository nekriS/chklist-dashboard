from PostLink import PostLinkClient

from system import log, create_thread_task, options
from functions import check_function, get_tmp_parts_from_db, update_status, start_event_checker, update_dashboard



if __name__ == "__main__":
    
    log("START")
    options = options()
    options.print_all_options()
    options.create_folders()

    client = PostLinkClient(options.config["BOT"]['API_BASE_URL'], options.config["BOT"]['WS_URL'], silent=False, event_checker=start_event_checker(options.config))
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

    #load_data(f"{config["GENERAL"]['DEFAULT_PATH']}{config["GENERAL"]['NAME_FOLDER_DATA']}/{config["GENERAL"]['NAME_FILE_DATA']}")
    #create_task_notification(options.config, data_all, project, target, type_n, step=3, count=356)

    while True:
        command = input("")
        log(command)
        req_massive = command.split(" ")
        match req_massive[0]:
            case "exit":
                destroy_check_function.set()
                destroy_tmp_parts.set()
                break
            case "reload":
                destroy_check_function.set()
                destroy_tmp_parts.set()
                options.update()
                options.print_all_options()
                thread, destroy_check_function = create_thread_task(options.config["GENERAL"]['CHECK_TIMEOUT'], check_function, options.config, client)
                thread_2, destroy_tmp_parts = create_thread_task(options.config["DB"]['CHECK_TIMEOUT'], get_tmp_parts_from_db, options.config)
            case "update":
                update_dashboard(options.config)
                update_status(options.config)
                update_dashboard(options.config)
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