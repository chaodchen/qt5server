from PyQt5.QtNetwork import QTcpSocket, QHostAddress

def get_local_ip():
    socket = QTcpSocket()
    socket.connectToHost(QHostAddress("8.8.8.8"), 80)  # 连接到一个远程地址
    if socket.waitForConnected():
        local_ip = socket.localAddress().toString()
        socket.close()
        return local_ip
    return ""


if __name__ == "__main__":
    local_ip = get_local_ip()
    print(local_ip)

