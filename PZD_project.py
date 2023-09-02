from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QFont,QDrag
from PyQt5 import QtCore
from PyQt5 import QtGui

import sys
from PyQt5.QtGui import QPalette, QBrush, QPixmap
import pandas as pd
import webbrowser
from openpyxl import load_workbook

count =0
cnt=0
coutn_window = False
coutn_window_2 = False
cnt1=0
cnt2=0
cnt3=0
cnt4=0
cnt5=0
cnt6=0
cnt7=0
cnt8=0
cnt9=0
gl3 = 0
file=''
ident_btn =''
txtn = ''
text_on_vivod = ''
global_strok = ''
global_input = ''
global_out = ''
global_strok_3 = False
repeat_words = ["повтор","не измен","СРОЧНО!!!","Вопрос не решен","Вопрос не решён"] # список повторных обращений

bad_words = ['хам', 'плох',"груб", "игнор","не отвечают на телефонные вызовы","не берут трубку",
              "не отвечают на сообщения", "нет обратной связи", "не взял", "скотина", "тварь", "сука","черт","невежа","грубиян"]
win_width, win_height = 500, 700
win_x, win_y = 800, 80
txt_title = "Приложение анализа качества предоставляемых услуг на основе обработки обращений."
stop = True

class SecondWindow(QWidget):
    
    def __init__(self, parent=None, flags=Qt.WindowFlags()):
        super().__init__(parent=parent, flags=flags) 
        global global_strok,global_input,global_out
        
        self.initUI()
        self.set_appear()  
        self.connects()
    def initUI(self):
        
        self.grid = QGridLayout()
        
        self.line1 = QTextEdit()

        self.line1.append(global_strok)
        self.line2 = QTextEdit(global_input)
        self.line2.setFixedSize(1000, 150)
        self.line22 = QLineEdit('')
        self.line22.setFixedSize(1000, 70)
        self.btn = QPushButton('Сохранить', self)
        self.grid.addWidget(self.line1, 1, 1)
        self.grid.addWidget(self.line2, 2,1)
        self.grid.addWidget(self.line22, 3,1)
        self.grid.addWidget(self.btn,4,1)

        self.setLayout(self.grid)
        QtCore.QMetaObject.connectSlotsByName(self)
  #Обновление экрана
    def shor(self):
        global global_strok,global_input
        self.line1.append(global_strok)
        self.line2.append(global_input)
        
    def set_appear(self):
        self.setWindowTitle('Окно вывода')
        self.resize(win_width+500, win_height)
        self.move(win_x-800, win_y)
    def connects(self):
        self.btn.clicked.connect(self.clicked)
    def clicked(self):
        global global_out,file,global_strok,global_input, ident_btn,coutn_window
        global_input = ''
        if coutn_window == False:
        
            self.show()
            coutn_window = True
        else:
            global_strok = ''
            self.close()
            self.show()
            coutn_window = True
        if ident_btn == '8.1':
            station = list(self.line22.text().split(', '))[0]
            df=pd.read_excel(file)
            global_input = ''
            res2=df.loc[df['Станция задержки вагонов'] == station]
            res2=res2.groupby(["Станция задержки вагонов","Подтема"])["Подтема"].count().reset_index(name="Количество")
            table2=res2.sort_values(by=["Количество"],ascending=False)
            global_strok = (f"таблица по станции: {station}") + '\n'
            self.shor()
            global_strok = table2.head(16).to_string()
            self.shor()
            global_strok ='\n'
            self.shor()
            road = list(self.line22.text().split(', '))[1].upper()
            unit = list(self.line22.text().split(', '))[2].upper()
            res3=df.loc[df["Дорога"]==road]
            res3=res3.groupby(["Дорога","Подтема"])["Подтема"].count().reset_index(name="Количество")
            table3=res3.sort_values(by=["Количество"],ascending=False)
            global_strok = (f"таблица по дороге: {road}") + '\n'
            self.shor()
            
            global_strok = table3.head(16).to_string()
            self.shor()
            res4=df.loc[df["Ответственный за решение"]==unit]
            res4=res4.groupby(["Ответственный за решение","Подтема"])["Подтема"].count().reset_index(name="Количество")
            table4=res4.sort_values(by=["Количество"],ascending=False)
            
            global_strok = (f"таблица по подразделению: {unit}") + '\n'
            self.shor()
            global_strok = table4.head(16).to_string()
            self.shor()
            global_strok = '\nстатистика по жд выполнена'
            self.shor()
            
            with pd.ExcelWriter("count_task8_2.xlsx") as writer:
                table2.to_excel(writer)
                writer.save()
            with pd.ExcelWriter("count_task8_3.xlsx") as writer:
                table3.to_excel(writer)
                writer.save()
            with pd.ExcelWriter("count_task8_4.xlsx") as writer:
                table4.to_excel(writer)
                writer.save()
            ident_btn = ''
        if ident_btn == '5':
            client=self.line22.text().upper()
            df=pd.read_excel(file, usecols=["Станция задержки вагонов","Наименование контакта","Подтема","Суть обращения"])
            df['Клиент']=df['Наименование контакта']
            df['Клиент'].fillna('Без названия', inplace=True)
            df["Клиент"]= df["Клиент"].astype('string')
            df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))
            df['repeat'] = df['Суть обращения'].apply(mn.check)
            rez=df.loc[df['repeat'] == True]
            rez_client=rez.loc[df['Клиент']==client]
            rez_rep=rez_client.groupby(['Клиент','Станция задержки вагонов'])['Станция задержки вагонов'].count().reset_index(name="count")
   #         with pd.ExcelWriter(f'таблица повтор обращений по {client}.xlsx', engine = 'openpyxl') as writer:
       #         rez_rep.to_excel(pd.ExcelWriter(f'таблица повтор обращений по {client}.xlsx'))
            global_strok = (f"таблица повтор обращений по клиенту {client}")
            self.shor()
            table5=rez_rep.sort_values('count', ascending=False)
            global_strok = table5.head(16).to_string()
            self.shor()
            global_strok = ''
            ident_btn = ''
        if ident_btn == '2':
            df=pd.read_excel(file)
            df["Станция задержки вагонов"]=df["Станция задержки вагонов"].str.upper()
            road=list(self.line22.text().split(', '))[0].upper()
            
            res1=df.loc[df["Дорога"]==road]
            res2=res1.groupby(["Дорога","Станция задержки вагонов"])["Станция задержки вагонов"].count().reset_index(name="Количество")
            res2=res2.sort_values("Количество",ascending=False)
            global_strok = (f"Таблица количества обращений по дороге {road}")
            self.shor()
            global_strok = (res2.head(10).to_string())
            self.shor()
            
            with pd.ExcelWriter(f"count_roads_{road}.xlsx") as writer:
                res2.to_excel(writer) 

            
            station1 = list(self.line22.text().split(', '))[1].upper()
            res3=df.loc[df["Станция задержки вагонов"]==station1]
            res4=res3.groupby(["Дорога","Станция задержки вагонов","Подтема"])["Подтема"].count().reset_index(name="Количество")
            res4=res4.sort_values("Количество",ascending=False)
            global_strok = (f"Таблица количества обращений по станции")
            self.shor()
            global_strok = (res4.head(16).to_string())
            self.shor()
            
         #   with pd.ExcelWriter("count_roads_2.xlsx") as writer:
        #        res4.to_excel(writer) 
            ident_btn = ''
        if ident_btn == '3':
            client_name=self.line22.text().upper()
            df=pd.read_excel(file)
            df["Клиент"]= df['Наименование контакта']
            df["Клиент"].fillna('Без названия', inplace=True)
            df["Клиент"]= df["Клиент"].astype('string')
            df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))

            res1=df.loc[df["Клиент"]==client_name]
            res2=res1.groupby(["Клиент","Подтема"])["Подтема"].count().reset_index(name="Количество")
            res3=res1.groupby(["Клиент","Ответственный за решение"])["Ответственный за решение"].count().reset_index(name="Количество")
            res2=res2.sort_values("Количество",ascending=False)
            res3=res3.sort_values("Количество",ascending=False)
            global_strok = (f"таблица {client_name} по подтемам")
            self.shor()
            global_strok = res2.head(16).to_string()
            self.shor()
            
            with pd.ExcelWriter("count_clients_themes.xlsx") as writer:
                res2.to_excel(writer)
            with pd.ExcelWriter("count_clients_units.xlsx") as writer:
                res3.to_excel(writer)
            ident_btn = ''
        if ident_btn == '4':
            df=pd.read_excel(file)
            df["Клиент"]= df['Наименование контакта']
            df["Клиент"].fillna('Без названия', inplace=True)
            df["Клиент"]= df["Клиент"].astype('string')
            df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))
            station_name=self.line22.text()

            res1=df.loc[df["Станция задержки вагонов"]==station_name]

            res2=res1.groupby(["Станция задержки вагонов","Подтема"])["Подтема"].count().reset_index(name="Количество")
            
            res2=res2.sort_values("Количество",ascending=False)
            global_strok = res2.head(16).to_string()

            with pd.ExcelWriter("count_stations_theme.xlsx") as writer:
                res2.to_excel(writer)
            self.shor()
            ident_btn = ''
class MainWindow(QWidget):
    def __init__(self, parent=None, flags=Qt.WindowFlags()):
        super().__init__(parent=parent, flags=flags)        
        self.initUI()
        self.connects()
        self.set_appear()
       # self.setStyleSheet("background-color: bisque;")
        self.show()
        
        
    def show_window_2(self):
        self.w2 = SecondWindow()
        self.w2.show()
       
    def initUI(self):
        self.layout_line = QVBoxLayout()
        self.layout_line1 = QHBoxLayout()
        self.layout_line2 = QHBoxLayout()
        self.layout_line3 = QHBoxLayout()
        self.layout_line4 = QHBoxLayout()
        self.layout_line5 = QHBoxLayout()
        self.layout_line6 = QHBoxLayout()    
        self.layout_line7 = QHBoxLayout()
        self.layout_line8 = QHBoxLayout()
        self.layout_line9 = QHBoxLayout()
        self.layout_line10 = QHBoxLayout()
        self.layout_line11 = QHBoxLayout()

#Кнопки
        self.btn_1 = QPushButton('Сохранить', self)
        self.btn_2str = QPushButton('1) Статистика по железным дорогам', self)
        self.btn_3str = QPushButton('2) Статистика по клиентам', self)
        self.btn_4str = QPushButton('3) Статистика по станциям задержки вагонов', self)
        self.btn_5str = QPushButton('4) Статистика, поступающим повторно', self)
        self.btn_6str = QPushButton('5) Фактам некорректного общения с клиентами \nсо стороны сотрудников ОАО «РЖД»', self)
        self.btn_7str = QPushButton('6) Объемам обращений, отнесенных на\nответственность различных структурных \nподразделений ОАО «РЖД»', self)
        self.btn_8str = QPushButton('7) %  обращений решенных без повторных обращений', self)
        self.btn_9str = QPushButton('8) Вывести статистику по проблемам', self)
        self.btn_10str = QPushButton('9) Вывести станции с максимальным числом задержек', self)

        self.btn_1.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_2str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_4str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_6str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_8str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_10str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_3str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_5str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_7str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        self.btn_9str.setStyleSheet('background: rgb(238, 59, 59);color: white; border: 1px solid black;')
        
#строки,добавление виджетов на строки
        self.line1 = QLabel('Введите файл:')
        self.line2 = QLineEdit('')

        self.line_bot = QPushButton('Ссылка на бота с инструкцией: https://t.me/RussianRailwaysQualityAppBot',self)
        self.line_bot.setStyleSheet('color: black; border: 2px solid black;')
        self.layout_line1.addWidget(self.line1)
        
        self.layout_line1.addWidget(self.line2)
        self.layout_line1.addWidget(self.btn_1)
        self.layout_line2.addWidget(self.btn_2str)
        self.layout_line3.addWidget(self.btn_3str)
        self.layout_line4.addWidget(self.btn_4str)
        self.layout_line5.addWidget(self.btn_5str)
        self.layout_line6.addWidget(self.btn_6str)
        self.layout_line7.addWidget(self.btn_7str)
        self.layout_line8.addWidget(self.btn_8str)
        self.layout_line9.addWidget(self.btn_9str)
        self.layout_line10.addWidget(self.btn_10str)
        self.layout_line11.addWidget(self.line_bot)
        self.layout_line.addLayout(self.layout_line1)
        self.layout_line.addLayout(self.layout_line2)
        self.layout_line.addLayout(self.layout_line3)
        self.layout_line.addLayout(self.layout_line4)
        self.layout_line.addLayout(self.layout_line5)
        self.layout_line.addLayout(self.layout_line6)
        self.layout_line.addLayout(self.layout_line7)
        self.layout_line.addLayout(self.layout_line8)
        self.layout_line.addLayout(self.layout_line9)
        self.layout_line.addLayout(self.layout_line10)
        self.layout_line.addLayout(self.layout_line11)
        

        
        # шрифт строк
        self.line1.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.line2.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_1.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_2str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_3str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_4str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_5str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_6str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_7str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_8str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_9str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.btn_10str.setFont(QFont('Arial Bold',12, QFont.Bold))
        self.line_bot.setFont(QFont('Arial Bold', 12, QFont.Bold))
        
        #добавление строк на главную
        self.setLayout(self.layout_line)
        # присоединение окон в общий обЪект
        QtCore.QMetaObject.connectSlotsByName(self)
        
     # функция проверки повторяющихся слов для повторных обращений
    def webbb(self):
        webbrowser.open('https://t.me/RussianRailwaysQualityAppBot', new=2)
    def check(self,string: str):
        global repeat_words
        string = str(string)
        string = string.lower()
        if string == '':
            return False
        for word in repeat_words:
            if word in string:
                return True
        return False
    def passinf(self):
        global stop
        while stop:
            pass
    def command_task7(self,x):
        global coutn_window, global_strok, ident_btn
       
        if coutn_window == False:
        
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        
        df=pd.read_excel(x,usecols=["Наименование контакта","Суть обращения"])

        count_texts=len(df)

        df["Клиент"]= df['Наименование контакта']
        df["Клиент"].fillna('Без названия', inplace=True)
        df["Клиент"]= df["Клиент"].astype('string')
        df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))

        df['repeat']=df['Суть обращения'].apply(self.check)
        res=df.loc[df['repeat']==True]
        count_all_rep=len(res)

        res=res.groupby(["Клиент"])["Клиент"].count().reset_index(name="Количество")
        res=res.sort_values('Количество', ascending=False)
        #res=res.loc[res['Количество']>1]
        striii = res.head(16).to_string()
        global_strok = striii + '\n'
        

        self.w2.shor()
        count_rep=len(res)
        
        static=round((count_texts-(count_all_rep+count_rep))*100/count_texts,3)
        global_strok = (f' % обращений, решенных без повторных обращений: {static}')
        self.w2.shor()
        
    def command_task8(self,x):
        global coutn_window, global_strok, global_input,ident_btn
        ident_btn = '8.1'
        
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        
        
        
        df=pd.read_excel(x)
        global_strok = "таблица по подтемам" + '\n'
        self.w2.shor()
        res=df.groupby(["Подтема"])["Подтема"].count().reset_index(name="Количество")
        res1=df.groupby(["Станция задержки вагонов"])["Станция задержки вагонов"].count().reset_index(name="Количество")

        table=res.sort_values(by=['Количество'], ascending=False)
        global_strok = table.head(16).to_string()
        self.w2.shor()

        table1=res1.sort_values(by=["Количество"],ascending=False)
        global_strok = "таблица по станциям" + '\n'
        self.w2.shor()
        global_strok = table1.head(16).to_string()
        self.w2.shor()
        global_strok = ''
        global_input = "Введите станцию:"+ '\n'
        self.w2.shor()
        global_input = "Введите дорогу:"+ '\n'
        self.w2.shor()
        global_input = "Введите подразделение РЖД:" + '\n'
        self.w2.shor()
        global_input = "Введите каждый элемент на одной строке через запятую с пробелом "
        self.w2.shor()
        global_input = ''
      #  self.t1.join()
     #   print(self.t1.is_alive())    
        with pd.ExcelWriter("count_task8.xlsx") as writer:
            table.to_excel(writer)
            writer.save()
        with pd.ExcelWriter("count_task8_1.xlsx") as writer:
            table1.to_excel(writer)
            writer.save()

    def command_task9(self,x):
        
        global coutn_window, global_strok,global_input
        global_input = ''
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        
        
        df=pd.read_excel(x,usecols=["Станция задержки вагонов","Подтема"])
        table=df.loc[df["Подтема"]=="Задержка вагонов в пути следования"]
        table=table.groupby(["Станция задержки вагонов","Подтема"])["Станция задержки вагонов"].count().reset_index(name="Количество")
        table=table.sort_values("Количество",ascending=False)

        with pd.ExcelWriter("count_task9.xlsx") as writer:
            table.to_excel(writer)
        global_strok = table.head(16).to_string()
        self.w2.shor()
        global_strok = str(len(table))
        self.w2.shor()
        

    def check_bad_word(self,string: str):
        global bad_words
        string = str(string)
        string = string.lower()
        if string == '':
            return False
        for word in bad_words:
            if word in string:
                return True
        return False

    #выполнение команды по грубым обращениям
    def command_rudeness(self,x):
        global coutn_window,global_strok,global_input
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        global_strok = ''
        
        df=pd.read_excel(x, usecols=["Станция задержки вагонов","Наименование контакта","Подтема","Суть обращения"])
        df['bad_words'] = df['Суть обращения'].apply(self.check_bad_word)

        rez=df.loc[df['bad_words'] == True]
        # уточнить проблему переменных
        rez1=rez.groupby(['Станция задержки вагонов'])['Станция задержки вагонов'].count().reset_index(name="count")
        rez2=rez.groupby(['Подтема'])['Подтема'].count().reset_index(name="count")
        global_strok = ("таблица грубых обращений относительно станций")
        self.w2.shor()
        global_strok = rez1.sort_values('count', ascending=False).head(16).to_string()
        self.w2.shor()
        global_strok ="таблица грубых обращений относительно подтемы"
        self.w2.shor()
        global_strok = rez2.sort_values('count', ascending=False).head(16).to_string()
        self.w2.shor()
        
        

        table1=rez1.sort_values('count', ascending=False)
        table2=rez2.sort_values('count', ascending=False)
        with pd.ExcelWriter("count_rudeness_stations.xlsx") as writer:
            table1.to_excel(writer)
        with pd.ExcelWriter("count_rudeness_themes.xlsx") as writer:
            table2.to_excel(writer)

    #выполнение команды по повторным обращениям
    def command_repeat(self,x):
        global coutn_window,global_strok,ident_btn,global_input
        ident_btn = '5'
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        global_input = "Введите клиента:"+ '\n'
        self.w2.shor()
        global_input = ''
        df=pd.read_excel(x, usecols=["Станция задержки вагонов","Наименование контакта","Подтема","Суть обращения"])

        df['Клиент']=df['Наименование контакта']
        df['Клиент'].fillna('Без названия', inplace=True)
        df["Клиент"]= df["Клиент"].astype('string')
        df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))


        df['repeat'] = df['Суть обращения'].apply(self.check)
        rez=df.loc[df['repeat'] == True]
        rez1=rez.groupby(['Станция задержки вагонов'])['Станция задержки вагонов'].count().reset_index(name="count")
        rez2=rez.groupby(['Клиент'])['Клиент'].count().reset_index(name="count")
        rez3=rez.groupby(['Подтема'])['Подтема'].count().reset_index(name="count")
        
        
        
        global_strok = '\n таблица повтор обращений'
        self.w2.shor()
        tabel1=rez
      #  global_strok = rez.sort_values('count', ascending=False).head(16).to_string()
      #  self.w2.shor()
        
        global_strok = '\n таблица повтор обращений относительно станций'
        self.w2.shor()
        
        tabel2=rez1.sort_values('count', ascending=False)
        global_strok = rez1.sort_values('count', ascending=False).head(16).to_string()
        self.w2.shor()

        global_strok = '\n таблица повтор обращений относительно контакта'
        self.w2.shor()
        
        tabel3=rez2.sort_values('count', ascending=False)
        global_strok = rez2.sort_values('count', ascending=False).head(16).to_string()
        self.w2.shor()
        
        global_strok = '\n таблица повтор обращений относительно подтемы'
        self.w2.shor()
        
        tabel4=rez3.sort_values('count', ascending=False)
        global_strok = rez3.sort_values('count', ascending=False).head(16).to_string()
        self.w2.shor()
    

        
        

        with pd.ExcelWriter("таблица повтор обращений.xlsx") as writer:
            tabel1.to_excel(writer)

        with pd.ExcelWriter("таблица повтор обращений относительно станций.xlsx") as writer:
            tabel2.to_excel(writer)

        with pd.ExcelWriter("Таблица повтор обращений относительно контакта.xlsx") as writer:
            tabel3.to_excel(writer)

        with pd.ExcelWriter("таблица повтор обращений относительно подтемы.xlsx") as writer:
            tabel4.to_excel(writer)

    #выполнение команды ЖД
    def command_road(self,x):
        global coutn_window,global_strok,ident_btn,global_input
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        ident_btn = '2'
        df=pd.read_excel(x)
        df["Станция задержки вагонов"]=df["Станция задержки вагонов"].str.upper()
        res=df.groupby(["Дорога"])["Дорога"].count().reset_index(name="Количество")
        res=res.sort_values("Количество",ascending=False)
        global_strok = res.head(10).to_string()
        global_input = "Введите дорогу:"+ '\n'
        global_input += ("Введите Станцию задержки вагонов:")
        global_input += '\n'
        global_input += 'Введите элементы через запятую с пробелом на одной строке'
        
        self.w2.shor()
        global_strok = ''
        
        global_input = ''
        with pd.ExcelWriter("count_road.xlsx") as writer:
            res.to_excel(writer) 

    # выполнение команды клиенты
    def command_clients(self,x):
        global coutn_window,global_strok,ident_btn,global_input
        ident_btn = '3'
        
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True

        df=pd.read_excel(x)

        global_input = "Введите Клиента:"+ '\n'
        self.w2.shor()
        global_input = ''
        df=pd.read_excel(x)
        df["Клиент"]= df['Наименование контакта']
        df["Клиент"].fillna('Без названия', inplace=True)
        df["Клиент"]= df["Клиент"].astype('string')
        df["Клиент"]= df["Клиент"].apply(lambda x: x.upper().replace("  "," ").replace('«','"').replace('»','"'))

        res=df.groupby(["Клиент"])["Клиент"].count().reset_index(name="Количество")
        res=res.sort_values("Количество",ascending=False)
        global_strok = ("Таблица обращений по клиентам")
        global_strok =(res.head(16).to_string()) 
        self.w2.shor()
        with pd.ExcelWriter("count_clients.xlsx") as writer:
            res.to_excel(writer)

        

    # выполнение команды станции

    def command_stations(self,x):
        global coutn_window,global_strok,global_input,ident_btn 
        ident_btn = '4'
        
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True

        df=pd.read_excel(x)
        df["Станция задержки вагонов"]=df["Станция задержки вагонов"].str.upper()
        res=df.groupby(["Станция задержки вагонов"])["Станция задержки вагонов"].count().reset_index(name="Количество")
        res=res.sort_values("Количество",ascending=False)
        global_strok = res.head(10).to_string()
        self.w2.shor()
        global_strok = ''
        with pd.ExcelWriter("count_stations.xlsx") as writer:
                res.to_excel(writer)
        global_input = "Введите Станцию:"+ '\n'
        self.w2.shor()
        global_input = ''






    # выполнение команды по подразделениям РЖД
    def command_units(self,x):
        global coutn_window,global_strok,global_input
        global_input = ''
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            global_strok = ''
            self.w2.close()
            self.show_window_2()
            coutn_window = True

        df=pd.read_excel(x, usecols=["Ответственный за решение"])
        #df.info()
        df['Name']=df['Ответственный за решение']
        #df.info()
        res=df.groupby(['Name']).count()
        res=res.rename(columns={"Ответственный за решение": "Количество"})
        res=res.sort_values(by=['Количество'], ascending=False)
        global_strok = res.head(16).to_string()
        self.w2.shor()
        
        with pd.ExcelWriter("count_units.xlsx") as writer:
            res.to_excel(writer)
    
        
    def roads(self):
        global cnt1,global_strok,file,global_input
        global_input = ''
        if cnt1==0:
            self.command_road(file)
            global_strok = '\nТаблица обращений по жд выполнена'
            self.w2.shor()
            global_strok = ''
            cnt1+=1
        else:
            global_strok = ('\nТаблица уже готова')
            self.w2.shor()
            global_strok = ''
            

        #обьем обращения по клиентам  +
    def clients(self):
        global cnt2,global_strok,file,global_input
        global_input = ''
        if cnt2==0:
            global_strok = ('\nТаблица обращений по клиентам выполнена')
            self.w2.shor()
            global_strok = ''
            self.command_clients(file)
            cnt2+=1
        else:
            global_strok =  '\nТаблица уже готова'
            self.w2.shor()
            global_strok = ''
            

    def task7(self):
        global cnt7,file,global_strok,global_input
        global_input = ''
        if cnt7==0:
            self.command_task7(file)
            cnt7+=1
        else:
            global_strok = '\n% уже выведен'
            self.w2.shor()
            global_strok = ''
            

    def task8(self):
        global cnt8,file,global_strok,global_input
        global_input = ''
        if cnt8==0:
            self.command_task8(file)
            
            cnt8+=1
        else:
            global_strok = ('\nстатистика уже готова')
            self.w2.shor()
            global_strok = ''

    def task9(self):
        global cnt9,global_strok,file,global_input
        global_input = ''
        if cnt9==0:
            self.command_task9(file)
            global_strok = ('\nТаблица обращений по станциям с макс числом обращений выполнена')
            self.w2.shor()
            global_strok = ''
            cnt9+=1
        else:
            global_strok = ('\nТаблица уже готова')
            
            self.w2.shor()
            global_strok = ''


#загрузка файла  +
    def clicked(self):
        global file, cnt, cnt1, cnt2, cnt3, cnt4, cnt5, cnt6, cnt7, cnt8, cnt9, coutn_window, coutn_window_2, global_strok_3, global_input, global_strok, text_on_vivod
        if coutn_window == False:
            self.show_window_2()
            coutn_window = True
        else:
            self.w2.close()
            self.show_window_2()
            coutn_window = True
        if cnt==0:
            
            file = "{}".format(self.line2.text())
            global_strok = ( f'\nФайл {file} загружен')
            self.w2.shor()
            cnt+=1
        else:
            if coutn_window == True:
                self.w2.close()
                coutn_window = False
            if coutn_window_2 == True:
                self.w3.close()
                coutn_window_2 = False
            
            cnt1=0
            cnt2=0
            cnt3=0
            cnt4=0
            cnt5=0
            cnt6=0
            cnt7=0
            cnt8=0
            cnt9=0
            file=''
            
            text_on_vivod = ''
            global_strok = ''
            global_input = ''
            global_strok_3 = ''
            coutn_window = False
            coutn_window_2 = False
            if coutn_window == False:
                self.show_window_2()
                coutn_window = True
            else:
                self.w2.close()
                self.show_window_2()
                coutn_window = True
            global_strok = ( '\nДанные обновлены')
            self.w2.shor()
            
            file = "{}".format(self.line2.text())
            global_strok = ( f'\nФайл {file} загружен')
            self.w2.shor()
            
    #обьем обращения по станциям +
    def stations(self):
        global cnt3,global_strok,file,global_input
        global_input = ''
        if cnt3==0:
            self.command_stations(file)
            global_strok = ( '\nТаблица обращений по станциям выполнена')
            self.w2.shor()
            global_strok = ''
            cnt3+=1
        else:
            global_strok = ( '\nТаблица уже готова')
            self.w2.shor()
            global_strok = ''
            

    #обьем обращения повторные
    def repeat(self):
        global cnt4,global_strok,file,global_input
        global_input = ''
        if cnt4==0:
            self.command_repeat(file)
            global_strok = ('\nТаблица обращений по повторным выполнена')
            self.w2.shor()
            global_strok = ''
            cnt4+=1
        else:
            global_strok = ( '\nТаблица уже готова')
            self.w2.shor()
            global_strok = ''
            
        

    #обьем обращения по грубости
    def rudeness(self):
        global cnt5,global_strok,file,global_input
        global_input = ''
        if cnt5==0:
            self.command_rudeness(file)
            global_strok = ('\nТаблица обращений по хамству выполнена')
            self.w2.shor()
            global_strok = ''
            cnt5+=1
        else:
            global_strok = ( '\nТаблица уже готова')
            self.w2.shor()
            global_strok = ''

    #обьем обращений по подразделениям РЖД  +
    def units(self):
        global cnt6,global_strok,file,global_input
        global_input = ''
        if cnt6==0:
            self.command_units(file)
            global_strok = ('\nТаблица обращений по подразделениям выполнена')
            self.w2.shor()
            global_strok = ''
            cnt6+=1
        else:
            global_strok = ( '\nТаблица уже готова')
            self.w2.shor()
            global_strok = ''
             
    def connects(self):
        self.btn_1.clicked.connect(self.clicked)
        self.btn_2str.clicked.connect(self.roads)
        self.btn_3str.clicked.connect(self.clients)
        self.btn_4str.clicked.connect(self.stations)
        self.btn_5str.clicked.connect(self.repeat)
        self.btn_6str.clicked.connect(self.rudeness)
        self.btn_7str.clicked.connect(self.units)
        self.btn_8str.clicked.connect(self.task7)
        self.btn_9str.clicked.connect(self.task8)
        self.btn_10str.clicked.connect(self.task9)
        self.line_bot.clicked.connect(self.webbb)
    
    
    def set_appear(self):
        self.setWindowTitle(txt_title)
        self.resize(win_width, win_height)
        self.move(win_x+300, win_y)
    

          
        
# создание окна
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('con_rchd.png'))

    mn = MainWindow()
    
    
    
    app.exec_()