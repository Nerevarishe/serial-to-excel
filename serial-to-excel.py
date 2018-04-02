import serial
import datetime
import time
import xlsxwriter


#ser = serial.Serial('COM3')
ser = serial.Serial(
port = "COM3",
baudrate = 9600,
bytesize = serial.EIGHTBITS, 
parity = serial.PARITY_NONE,
stopbits = serial.STOPBITS_ONE, 
timeout = 1,
xonxoff = False,
rtscts = True,
dsrdtr = True,
writeTimeout = 2
)


print("connected to: " + ser.portstr)
#ser.open()
print(ser.isOpen())
#ser.close()
print(ser.isOpen())
#print(ser.read())

# this will store the line
seq = []
count = 1

# Создание новой excel книги. В качестве имени книги используется дата и время создания
wbName = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + '.xlsx'
wb = xlsxwriter.Workbook(wbName)
ws = wb.add_worksheet()

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
        
        if ser.isOpen == False:
            wb.close()
            break
