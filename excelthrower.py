  import socket
import time
from openpyxl import load_workbook
while True:
    IP = input('ql ip osoite > ')
    if not  isinstance(IP, str):
        print("anna validi IP osoite")
    else:
        break
tyokirja = input('excel tiedoston nimi > ')
wb = load_workbook(filename = tyokirja)

    
sheet_ranges = wb.active

kannu = []
kannunimi = []


TCP_PORT = '49280'

    
TCP_IP = IP

def main():
     for i in range(3,58):
              
         
         numero = sheet_ranges[f'C{i}']
        
         listaan = isinstance(numero.value, float)
         if listaan:
             kannu.append(int(numero.value))
     for i in range(3,58):

          nimi = sheet_ranges[f'E{i}']
        
          toiseenlistaan = isinstance(nimi.value, str)
          if toiseenlistaan:
              kannunimi.append(nimi.value)
     print(kannu)

     print()
     print(kannunimi) 

     TCP_IP = IP
     TCP_PORT = 49280
     s=socket.socket(socket.AF_INET, socket.SOCK_STREAM)
     s.connect((TCP_IP, TCP_PORT))

     for i in range(0, len(kannu)):
         cmd=(f'set MIXER:Current/InCh/Label/Name {kannu[i]-1} 0 "{kannunimi[i]}" "{kannunimi[i]}" ')
         send_cmd(s,cmd)
         cmd=(f'set MIXER:Current/InCh/Fader/Level {kannu[i]-1} 0 0 "0.00"')
         send_cmd(s,cmd)

def send_cmd(s, cmd):
    print(cmd)
    s.send(cmd.encode() +b'\n')
    time.sleep(0.01)
    print(s.recv(2048))


if __name__== '__main__':
    main()
