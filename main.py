from system import log, create_thread_task
from core.dashboard import dashboardObject
from core.services import check_function, get_tmp_parts_from_db
from web.server import serverStart


VERSION = "0.0.2.005 2026.10.22"

if __name__ == "__main__":

    log("Start point")

    dashboard = dashboardObject()
    dashboard.version = VERSION
    if dashboard.messanger.is_auth:
        while not dashboard.messanger.is_connected:
            pass
        else:
            dashboard.messanger.subscribe("/topic/messages")

    loop_timeout = dashboard.config["GENERAL"]['CHECK_TIMEOUT']
    db_fetcher_timeout = dashboard.config["DB"]['CHECK_TIMEOUT']
    port = int(dashboard.config["WEB"]['PORT'])

    thread_loop,        destroy_loop        = create_thread_task(loop_timeout, 
                                                                 check_function, 
                                                                 dashboard)
    thread_db_fetcher,  destroy_db_fetcher  = create_thread_task(db_fetcher_timeout, 
                                                                 get_tmp_parts_from_db, 
                                                                 dashboard.config)
    thread_web_server,  destroy_web_server  = serverStart(port = port, 
                                                          debug = False, 
                                                          logger = log)

    while True:
        command = input("")
        log(command)
        req_massive = command.split(" ")
        match req_massive[0]:
            case "exit":
                destroy_loop.set()
                destroy_db_fetcher.set()
                destroy_web_server()
                break