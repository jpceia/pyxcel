import socket
import threading
import json
import time
import sys


# hostname = socket.gethostname()
# host = socket.gethostbyname(hostname)
# priv_addr = s.getsockname()


class Server:
    connections = []
    peers = []
    def __init__(self, port=10000):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.bind(("0.0.0.0", port))
        s.listen(1)
        while True:
            c, a = s.accept()
            cThread = threading.Thread(target=self.handler, args=(c, a))
            cThread.daemon = True
            cThread.start()
            self.connections.append(c)
            self.peers.append(a[0])
            print(str(a[0]) + ":" + str(a[1]), "connected")
            self.sendPeers()

    def handler(self, c, a):
        while True:
            data = c.recv(1024)
            for conn in self.connections:
                conn.send(data)
            if not data:
                print(str(a[0]) + ":" + str(a[1]), "disconnected")
                self.connections.remove(c)
                self.peers.remove(a[0])
                c.close()
                self.sendPeers()
                break

    def sendPeers(self):
        for conn in self.connections:
            connection.send(b'\x11' + bytes(",".join(self.peers), "utf-8"))


class Client:

    def sendMsg(self):
        while True:
            self.sock.send(bytes(input(""), "utf-8"))

    def __init__(self, address, port=10000):
        self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.sock.connect((address, port))

        iThread = threading.Thread(target=self.sendMsg)
        iThread.daemon = True
        iThread.start()

        while True:
            data = self.sock.recv(1024)
            if not data:
                break
            if data[:1] == b'\x11':
                self.updatePeers(data[1:])
            else:
                print(str(data, 'utf-8'))

    def updatePeers(self, peerData):
        p2p.peers = str(peerData, "utf-8").split(",")


class p2p:
    peers = ['127.0.0.1']



def pscan(addr, port):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.connect((addr, port))
        return True
    except:
        return False

def scan_address(addr, ports=range(1, 81)):
    for p in ports:
        if pscan(addr, p):
            print("Port ", p, " is open")
        else:
            print("Port ", p, " is closed")

if __name__ == "__main__":
    while True:
        try:
            print("Trying to connect ...")
            time.sleep(1)
            for peer in p2p.peers:
                try:
                    client = Client(peer)
                except KeyboardInterrupt:
                    sys.exit(0)
                except:
                    pass

                try:
                    server = Server()
                except KeyboardInterrupt:
                    sys.exit(0)
                except:
                    print("Couldn't start the server ...")

        except KeyboardInterrupt:
            sys.exit(0)
