import serial
import time

ser = serial.Serial(
    port='COM3',\
    baudrate=9600,\
    parity=serial.PARITY_NONE,\
    stopbits=serial.STOPBITS_ONE,\
    bytesize=serial.EIGHTBITS,\
        timeout=0)

print("connected to: " + ser.portstr)

#this will store the line
seq = []
count = 1

while True:
    for c in ser.read():
        seq.append(chr(c)) #convert from ANSII
        joined_seq = ''.join(str(v) for v in seq) #Make a string from array

        if chr(c) == '\n':
            print("Line " + str(count) + ': ' + joined_seq)
            seq = []
            count += 1
            break


ser.close()
