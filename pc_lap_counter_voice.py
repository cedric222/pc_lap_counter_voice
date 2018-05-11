import socket
import re
import win32com.client as wincl

HOST = '127.0.0.1'  # The remote host
PORT = 7228  # The same port as used by the server

def read_lap():
    """Coroutine Reading Lap Time"""
    speak = wincl.Dispatch("SAPI.SpVoice")
    while True:
        pilote, lap_time  = yield
        if lap_time:
            lap_time = re.match(r"((?P<m>\d+):)?(?P<s>\d+)(\.(?P<ms>\d+))?", lap_time)
            if lap_time:
                to_speak = f"{pilote} "
                if lap_time.group('m'):
                    to_speak += (f"{lap_time.group('m')} minutes ")
                to_speak += (f"{lap_time.group('s')} secondes ")
                if lap_time.group('ms'):
                    to_speak += (f"{lap_time.group('ms')}")
                print(to_speak)
                speak.Speak(to_speak)


if __name__ == '__main__':
    read = read_lap()
    next(read)
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect((HOST, PORT))
        server_data = {}
        while True:
            try:
                data = s.recv(4096)
                for record in data.decode().split("$"):
                    if record:
                        try:
                            key, value = record.split("=")
                        except ValueError:
                            pass
                        else:
                            m = re.match(r"IO(?P<pilote_id>\d+)laptime", key)
                            if m and server_data.get(key) != value and value != "0.000":
                                pilote_name_key = f"IO{m.group('pilote_id')}Pilote"
                                read.send((server_data.get(pilote_name_key),
                                         value))
                            server_data[key] = value
            except Exception as e:
                print(e)
                pass
