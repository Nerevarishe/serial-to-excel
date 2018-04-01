import serial
import datetime
import time
import xlsxwriter


ser = serial.Serial(
    port='COM3',\
    baudrate=9600,\
    parity=serial.PARITY_NONE,\
    stopbits=serial.STOPBITS_ONE,\
    bytesize=serial.EIGHTBITS,\
        timeout=0)

#print("connected to: " + ser.portstr)

# this will store the line
seq = []
count = 1

# Создание новой excel книги. В качестве имени книги используется дата и время создания
wbName = datetime.datetime.now().strftime("%Y-%m-%d-%H:%M:%S") + '.xlsx'
wb = xlsxwriter.Workbook(wbName)
ws = workbook.add_worksheet()

# Выставляем курсор в первую ячейку (A1)
row = 0
col = 0


while True:
    for c in ser.read():
        seq.append(chr(c)) #convert from ASCII
        joined_seq = ''.join(str(v) for v in seq) #Make a string from array

        if chr(c) == '\n':
            ws.write(row, col, joined_seq)
            row += 1
            seq = []
            count += 1
            break


ser.close()
