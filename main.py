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
        self.connect.setEnabled(True)
        self.disconnect.setEnabled(False)
        self.statusBar.showMessage('Disconnected')
        
    # Отображение диалогового окна при ошибке:
    def onConnectionError(self, error):
        error_text = ''.join(map(str, error))
        error_dialog = QMessageBox()
        error_dialog.setText(error_text)
        error_dialog.exec()
        self.connect.setEnabled(True)
        self.disconnect.setEnabled(False)

    # Основная функция программы выполняющая все измерения 
    def workInThread(self):
        ser = serial.Serial(
                port = self.serialPort.text(),
                baudrate = self.baudRate.text(),
                timeout = 1,
                xonxoff = self.xonxoff.isChecked(),
                rtscts = self.rtscts.isChecked(),
                dsrdtr = self.dsrdtr.isChecked(),
                writeTimeout = 2
            )
        # КОСТЫЛЬ. Настройка комбо боксов. Разобраться потом 
        # как правильно отправлять данные.
        # stopBits
        if self.stopBits.currentText() == "serial.stopBits_ONE":
            ser.stopBits = 1

        elif self.stopBits.currentText() == "serial.stopBits_ONE_POINT_FIVE":
            ser.stopBits = 1.5
            
        elif self.stopBits.currentText() == "serial.stopBits_TWO":
            ser.stopBits = 2

        # byteSize
        if self.byteSize.currentText() == "serial.EIGHTBITS":
            ser.bytesize = 8

        elif self.byteSize.currentText() == "serial.SEVENBITS":
            ser.bytesize = 7

        elif self.byteSize.currentText() == "serial.SIXBITS":
            ser.bytesize = 6

        elif  self.byteSize.currentText() == "serial.FIVEBITS":
            ser.bytesize = 5

        # parity
        if self.parity.currentText() == "serial.PARITY_NONE":
            ser.parity = 'N'

        elif self.parity.currentText() == "serial.PARITY_EVEN":
            ser.parity = 'E'

        elif self.parity.currentText() == "serial.PARITY_ODD":
            ser.parity = 'O'

        elif  self.parity.currentText() == "serial.PARITY_MARK":
            ser.parity = 'M'

        elif  self.parity.currentText() == "serial.PARITY_SPACE":
            ser.parity = 'S'

        if ser.isOpen() == True:
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

        while True:
            for c in ser.read():
                #convert from ASCII
                seq.append(chr(c))
                joined_seq = ''.join(str(v) for v in seq) #Make a string from array

                # Если в серийном порту получен сигнал 'NEW_MEASUREMENT\r\n', перейти на следующую строку в экселе и обнулить переменную seq
                if joined_seq == 'NEW_MEASUREMENT\r\n':
                    row += 1
                    seq = []
                    break

                elif chr(c) == '\n':
                    ws.cell(column = col, row = row, value = joined_seq)
                    seq = []
                    #count += 1
                    wb.save(wbName)
                    break

                if ser.isOpen == False:
                    wb.save(wbName)
                    break


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