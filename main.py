from PyQt5 import QtGui, QtWidgets, QtCore

from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

import datetime
import traceback, sys
import serial
from openpyxl import Workbook

import gui


class WorkerSignals(QObject):
    error = pyqtSignal(tuple)


class Worker(QRunnable):
    def __init__(self, fn, *args, **kwargs):
        super(Worker, self).__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
    
    @pyqtSlot()
    def run(self):
        try:
            self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            pass
        finally:
            pass


class SerialToExcel(QtWidgets.QMainWindow, gui.Ui_MainWindow):
    def __init__(self):
        # Это здесь нужно для доступа к переменным, методам
        # и т.д. в файле gui.py
        super().__init__()
        # Это нужно для инициализации нашего дизайна
        self.setupUi(self)
        # Создание пула потоков
        self.threadpool = QThreadPool()
        
        """# Создание объекта серийного порта
        self.ser = seial.Serial()"""
        
        # Активация кнопки Connect
        self.disconnect.setEnabled(False)
        # Надпись в статус баре
        self.statusBar.showMessage('Disconnected')
        # Запуск чтения серийного порта при нажатии на кнопку Connect
        self.connect.clicked.connect(self.connectToSerial)
        # Остановка чтения и закрытие серийного порта, сохранение файла эксель.
        self.disconnect.clicked.connect(self.disconnectFromSerial)

    # Подключение к серийному порту после нажатия кнопки connect
    def connectToSerial(self):
        # Активация кнопки Disconnect и деактивация Connect
        self.disconnect.setEnabled(True)
        self.connect.setEnabled(False)
        print('Connect button pressed!')
        # Запуск измерения в отдельном потоке
        worker = Worker(self.workInThread)
        worker.signals.error.connect(self.onConnectionError)
        self.threadpool.start(worker)
    
    # Отключение от серийного порта после нажатия кнопки disconnect
    def disconnectFromSerial(self):
        print('Disconnect button pressed!')
        self.ser.close()
        self.connect.setEnabled(True)
        self.disconnect.setEnabled(False)
        self.statusBar.showMessage('Disconnected')
        
    # Отображение диалогового окна при ошибке:
    def onConnectionError(self, error):
        self.ser.close()
        error_text = ''.join(map(str, error))
        error_dialog = QMessageBox()
        error_dialog.setText(error_text)
        error_dialog.exec()
        self.connect.setEnabled(True)
        self.disconnect.setEnabled(False)

    # Основная функция программы выполняющая все измерения 
    def workInThread(self):
        self.ser = serial.Serial(
                port = self.serialPort.text(),
                baudrate = self.baudRate.text(),
                timeout = 1,
                xonxoff = self.xonxoff.isChecked(),
                rtscts = self.rtscts.isChecked(),
                dsrdtr = self.dsrdtr.isChecked(),
                writeTimeout = 2
            )
        """self.ser.port = self.serialPort.text()
        self.ser.baudrate = self.baudRate.text()
        self.ser.timeout = 1
        self.ser.xonxoff = self.xonxoff.isChecked()
        self.ser.rtscts = self.rtscts.isChecked()
        self.ser.dsrdtr = self.dsrdtr.isChecked()
        self.ser.writeTimeout = 2"""
        # КОСТЫЛЬ. Настройка комбо боксов. Разобраться потом 
        # как правильно отправлять данные.
        # stopBits
        if self.stopBits.currentText() == "serial.stopBits_ONE":
            self.ser.stopBits = 1

        elif self.stopBits.currentText() == "serial.stopBits_ONE_POINT_FIVE":
            self.ser.stopBits = 1.5
            
        elif self.stopBits.currentText() == "serial.stopBits_TWO":
            self.ser.stopBits = 2

        # byteSize
        if self.byteSize.currentText() == "serial.EIGHTBITS":
            self.ser.bytesize = 8

        elif self.byteSize.currentText() == "serial.SEVENBITS":
            self.ser.bytesize = 7

        elif self.byteSize.currentText() == "serial.SIXBITS":
            self.ser.bytesize = 6

        elif  self.byteSize.currentText() == "serial.FIVEBITS":
            self.ser.bytesize = 5

        # parity
        if self.parity.currentText() == "serial.PARITY_NONE":
            self.ser.parity = 'N'

        elif self.parity.currentText() == "serial.PARITY_EVEN":
            self.ser.parity = 'E'

        elif self.parity.currentText() == "serial.PARITY_ODD":
            self.ser.parity = 'O'

        elif  self.parity.currentText() == "serial.PARITY_MARK":
            self.ser.parity = 'M'

        elif  self.parity.currentText() == "serial.PARITY_SPACE":
            self.ser.parity = 'S'

        if self.ser.isOpen() == True:
            self.statusBar.showMessage('Connected to: ' + self.serialPort.text())

        else:
            self.statusBar.showMessage('Can\'t connect to: ' + self.serialPort.text())

        # this will store the line
        seq = []
        count = 1

        # Создание новой excel книги. В качестве имени книги используется дата и время создания
        wbName = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + '.xlsx'
        wb = Workbook()
        ws = wb.active

        wb.save(wbName)

        # Выставляем курсор в первую ячейку (A1)
        row = 1
        col = 1

        iterration = 0
        joined_seq = ''
        c = 0
        
        while self.ser.isOpen() == True:
            iterration += 1
            print('while loop iterration ' + str(iterration))
            try:
                if self.ser.isOpen() == True:
                    for c in self.ser.read():
                        #convert from ASCII
                        seq.append(chr(c))
                        joined_seq = ''.join(str(v) for v in seq) #Make a string from array
            
            except Exception:
                #pass
                traceback.print_exc()
            
            else:
                # Если в серийном порту получен сигнал 'NEW_MEASUREMENT\r\n', перейти на следующую строку в экселе и обнулить переменную seq
                if joined_seq == 'NEW_MEASUREMENT':
                    row += 1
                    seq = []

                elif chr(c) == '\n':
                    ws.cell(column = col, row = row, value = joined_seq)
                    print(joined_seq)
                    seq = []
                    c = 0
                    #count += 1
                    wb.save(wbName)
        wb.save(wbName)


def main():
    # Новый экземпляр QApplication
    app = QtWidgets.QApplication(sys.argv)
    # Создаём объект класса SerialToExcel
    window = SerialToExcel()
    # Показываем окно
    window.show()
    # Запускаем приложение
    app.exec_()

# Если мы запускаем файл напрямую, а не импортируем, то запускаем функцию main()
if __name__ == '__main__':
    main()