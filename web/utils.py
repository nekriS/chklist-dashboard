import socket

def get_computer_link(port=2000):
    hostname = socket.gethostname()
    return f"http://{hostname}:{port}"