import requests
import datetime
import websocket
import threading
import time
import json
import urllib.parse
from pathlib import Path
import uuid
import io

class LinkClient:
    def __init__(self, api_base_url, ws_url, silent=False, listener=None, logger=None):
        self.api_base_url = api_base_url
        self.ws_url = ws_url
        self.ws = None
        self.is_connected = False
        self.download_folder = "download"
        self.session = requests.Session()
        self.listener = listener
        self.silent = silent
        self.is_auth = False
        self.logger = logger
        self.debug = None

    def _log(self, message):
        if not self.silent:
            if self.logger == None:
                print(f"[L] [{datetime.datetime.now()}] > {message}")
            else:
                self.logger(f"[L] {message}")


    def _is_authorized(self):
        try:
            resp = self.session.get(f"{self.api_base_url}/auth", timeout=30)
            return resp.status_code == 200 and resp.text.strip().lower() == "true"
        except Exception as e:
            self._log(f"[AUTH ERROR]: {e}")
            return False
        
    def _on_ws_error(self, ws, error):
        self._log(f"[WS ERROR]: {error}")
    
    def _on_ws_close(self, ws, close_status_code, close_msg):
        self.is_connected = False
        self._log("[WS] The connection was closed.")
        self.listener(self, "CONNECTION_CLOSED", None)

    def _on_ws_open(self, ws):
        self._log("[WS] Connected. Sending STOMP CONNECT...")
        try:
            connect_frame = "CONNECT\naccept-version:1.2\nhost:localhost\n\n\x00"
            ws.send(connect_frame)
            self._log("[WS] STOMP CONNECT sent")
        except Exception as e:
            self._log(f"[WS OPEN ERROR] {e}")

    # Подписываемся на топики
    def subscribe(self, sub="/topic/messages"):
        if self.debug != "console":
            if self.is_connected:
                sub_frame = f"SUBSCRIBE\nid:1\ndestination:{sub}\n\n\x00"
                self.ws.send(sub_frame)
                self._log(f"[WS] Subscription to {sub} sent")
            else:
                self._log("[WS] Connection not established! Subscription not possible.")
        else:
            self._log("[WS] Running in console mode!")

    def _on_ws_message(self, ws, message):
        # STOMP-фреймы — это текст, завершённый \x00
        if isinstance(message, bytes):
            message = message.decode('utf-8')
        
        # Убираем завершающий нулевой байт, если есть
        if message.endswith('\x00'):
            message = message[:-1]

        # Если сообщение — это STOMP-фрейм CONNECTED, просто логируем
        if message.startswith("CONNECTED"):
            self._log("[WS] STOMP connection established")
            self.is_connected = True
            return

        # Если сообщение начинается с "MESSAGE", извлекаем тело
        if message.startswith("MESSAGE"):
            try:
                # Ищем пустую строку — после неё идёт тело
                header_end = message.find("\n\n")
                if header_end == -1:
                    self._log(f"[STOMP] Message body not found: {message}")
                    return
                json_body = message[header_end + 2:]
                data = json.loads(json_body)

                #print(data)
                #print(type(data))

                if isinstance(data, dict):
                    event_type = data.get("type")



                    self._log(f"Event: {event_type}")
                    payload = data.get("data", [])

                    for data_message in payload:

                        if isinstance(data_message, list):
                            data_message_ = data_message[0]
                        else:
                            data_message_ = data_message
                        
                        match event_type:
                            case "INCOMING_MESSAGE":
                                self._log(f"Input message from {data_message_.get('ownerId')}: {data_message_.get('text').split('<span')[0]}")

                                server_msg_id = data_message_["serverMessageId"]

                                if data_message_.get('hasFile', False):
                                    self._log(f"Files found in the message: {data_message_["serverFileIds"]}")
                                    for file_id in data_message_["serverFileIds"]:
                                        try:
                                            self.request_file_download(server_msg_id, file_id)
                                            
                                            self.download_file(
                                                session=self.session,
                                                base_url=self.api_base_url,
                                                server_msg_id=server_msg_id,
                                                server_file_id=file_id,
                                                save_dir=f"./{self.download_folder}"
                                            )

                                            if self.listener != None:
                                                self.listener(self, "FILE_RECEIVED", data_message_)
                                        except Exception as e:
                                            self._log(f"File download error {file_id}: {e}")

                                if self.listener != None:
                                    self.listener(self, event_type, data_message_)
                            case "NEW_MESSAGE":
                                self._log(f"Output message to {data_message_.get('targetId')}: {data_message_.get('text').split('<span')[0]}")
                                if self.listener != None:
                                    self.listener(self, event_type, data_message_)
                            case _:
                                pass
                                #self._log(f"Unknown event!")
            
            except Exception as e:
                self._log(f"[PARSE ERROR]: {e}")
        
        else:
            self._log(f"[STOMP FRAME]: {message}")

    def connect_websocket(self, debug=False):
        if self.debug != "console":
            websocket.enableTrace(debug)  # отключить debug-логи
            self.ws = websocket.WebSocketApp(
                self.ws_url,
                on_open=self._on_ws_open,
                on_message=self._on_ws_message,
                on_error=self._on_ws_error,
                on_close=self._on_ws_close
            )
            # Запускаем в фоновом потоке
            self.ws_thread = threading.Thread(target=self.ws.run_forever)
            self.ws_thread.daemon = True
            self.ws_thread.start()
        else:
            self.is_connected = True
        # Ждём подключения
        #time.sleep(1)

    def ensure_connected(self):
        if self.debug != "console":
            if not self._is_authorized():
                self.is_auth = False
                raise RuntimeError("Not authorized!")
            else:
                self.is_auth = True
        else:
            self._log("CONSOLE MODE!")
            self.is_auth = True
            
        self.connect_websocket()
        self._log("REST ans WebSocket are ready!")

    def send_message(self, target_id, target_type, text):
        payload = {
            "targetId": target_id,
            "targetType": target_type,
            "message": text
        }
        if self.debug != "console":
            resp = self.session.put(f"{self.api_base_url}/message", json=payload)
            resp.raise_for_status()
            return resp.json()
        else:
            self._log("CONSOLE MODE!")
            self._log(f"{payload}")
    
    def request_file_download(self, server_msg_id, server_file_id, i = 0):
        time.sleep(5)
        url = f"{self.api_base_url}/message/download-task"
        payload = {
            "serverMsgId": server_msg_id,
            "serverFileId": server_file_id
        }
        try:
            resp = self.session.post(url, json=payload)
            if resp.status_code == 201:
                self._log(f"File download task requested {server_file_id}")
                return True
            else:
                self._log(f"Task request error: HTTP {resp.status_code}")
                if i < 5:
                    self._log(f"Try again in 5 seconds.")
                    
                    self.request_file_download(server_msg_id, server_file_id, i + 1)
                else:
                    self._log(f"Failed to send request.")       

        except Exception as e:
            self._log(f"Exception while requesting task: {e}")
            return False
        
    def download_file(self, session, base_url, server_msg_id, server_file_id, save_dir="."):
        time.sleep(5)
        url = f"{base_url}/message/download"
        payload = {
            "serverMsgId": server_msg_id,
            "serverFileId": server_file_id
        }
        resp = session.post(url, json=payload)
        resp.raise_for_status()

        #print("here")
        #print(resp)

        # Имя файла из заголовка Content-Disposition
        #print(resp.headers)
        #print(resp.headers["Content-Disposition"])
        #print(resp.headers.Content-Disposition)
        content_disp = resp.headers.get("Content-Disposition", "")

        #print(content_disp)
        if "filename*" in content_disp:
            # Обработка URL-encoded имени (часто для кириллицы)
            filename = content_disp.split("filename*=UTF-8''")[-1]
            filename = urllib.parse.unquote(filename)
        elif "filename=" in content_disp:
            filename = content_disp.split("filename=")[-1].strip('"')
        else:
            filename = f"file_{server_file_id}.bin"

        filepath = Path(save_dir) / filename
        filepath.parent.mkdir(parents=True, exist_ok=True)

        with open(filepath, "wb") as f:
            f.write(resp.content)
        self._log(f"File saved: {filepath}")
        return filepath
    
    def send_file(self, target_id, filepath, target_type="person", message=""):

        
        with open(filepath, "rb") as f:
            file_data = f.read()

        if filepath.rfind("/") != -1:
            file_name = filepath[filepath.rfind("/"):]
        else:
            file_name = filepath
        #print(file_name)
        file_uuid = str(uuid.uuid4())
        msg_dto = {
            "targetId": target_id,
            "targetType": target_type,
            "message": message,
            "files": [{
                "fileUuid": file_uuid,
                "name": file_name,
                "length": len(file_data),
                "directory": False
            }]
        }

        if self.debug != "console":
            resp = self.session.put(f"{self.api_base_url}/message", json=msg_dto)
            resp.raise_for_status()
            msg_uuid = resp.json()["uuid"]

            resp = self.session.post(
                f"{self.api_base_url}/message/file",
                params={"msgUuid": msg_uuid, "fileUuid": file_uuid},
                headers={
                    "Upload-Offset": "0",
                    "Content-Length": str(len(file_data))
                },
                data=file_data
            )
            resp.raise_for_status()

            self._log("File was sended.")
        else:
            self._log("CONSOLE MODE!")
            self._log(f"{msg_dto}")

    def set_avatar(self, image):

        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        avatar_data = buffer.getvalue()

        resp = requests.patch(
            f"{self.api_base_url}/user/avatar",
            data=avatar_data,
            headers={"Content-Type": "application/octet-stream"}
        )

        if resp.status_code == 200:
            self._log("Avatar is set!")
        else:
            self._log(f"Error with avatar setting: {resp.status_code} — {resp.text}")