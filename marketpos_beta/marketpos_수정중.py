from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import pyqtSlot, QTimer, QEventLoop
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import QApplication, QWidget, QListWidget, QListWidgetItem, QLabel, QMainWindow, QTableWidgetItem, QDialog
from PyQt5 import uic
import sys

import openpyxl
from openpyxl.styles import Border, Side, Alignment

import csv
from datetime import datetime

import math

#변수 정의

file_directory = "test.xlsx" #마켓 엑셀 경로 복붙
file_directory2 = "kim.xlsx"
wb = openpyxl.load_workbook(file_directory) 
wb2 = openpyxl.load_workbook(file_directory2)
sheet = wb["이용자"]
sheet2 = wb2["Sheet1"]
name_col = "B" #이용자 이름에 해당되는 열
ssn_col = "D" #이용자 주민등록번호에 해당되는 열

expire_date = "L"
visit_time = "Q"
visit_date = "R"
delegate = "P"
memo = "O"
expiry_date = "L"
phone = "H"

p1_name = "S"
p1_quantity = "T"
p2_name = "U"
p2_quantity = "V"
p3_name = "W"
p3_quantity = "X"
p4_name = "Y"
p4_quantity = "Z"
p5_name = "AA"
p5_quantity = "AB"

sh2_name = "A"
sh2_supplier = "B"
sh2_poommok = "C"
sh2_quantity = "D"
sh2_barcode = "E"

thin_border = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
    )


Nlist = [p1_name, p2_name, p3_name, p4_name, p5_name]
Qlist = [p1_quantity, p2_quantity, p3_quantity, p4_quantity, p5_quantity]

lookup_dict = {}
product_list = []

FirstWindow_formclass = uic.loadUiType("FirstWindow.ui")[0]
SecondWindow_formclass = uic.loadUiType("SecondWindow.ui")[0]
ThirdWindow_formclass = uic.loadUiType("ThirdWindow.ui")[0]

#함수 정의

def create_dict():
    for ssn in sheet[ssn_col]:
        if ssn.value:                 #None까지 딕셔너리에 key로 들어가지 않게 하기 위함
            key_row = str(ssn.row)
            number = str(sheet[phone + key_row].value).split("-")
            pure_number = "".join(number)
            lookup_dict[str(ssn.value)] = [
                sheet[name_col + key_row].value,
                sheet[visit_date + key_row].value,
                sheet[delegate + key_row].value,
                sheet[expiry_date + key_row].value,
                pure_number
            ]

def create_plist():  

    for k in sheet2[sh2_name]:
        if k.value:                   #None까지 딕셔너리에 key로 들어가지 않게 하기 위함
            product_list.append(Product(str(k.row)))

def window():
    app = QApplication(sys.argv)
    win = FirstWindow()
    win.move(370,100)
    win.show()
    sys.exit(app.exec_())

#오브젝트 생성
# class User:
#     def __init__(self, lst):
#         self.name = lst[0]
#         self.visit = lst[1]
#         self.delegate = lst[2]
#         self.expiry
#         self.phone
#         self.adress

class Product:	

    def __init__(self, row):
        self.name = sheet2[sh2_name + row].value
        self.supplier = sheet2[sh2_supplier + row].value
        self.poommok = float(sheet2[sh2_poommok + row].value)
        self.quantity =	int(sheet2[sh2_quantity + row].value)
        self.multiply = math.ceil(self.poommok*self.quantity)
        self.barcode = str(sheet2[sh2_barcode + row].value)
        self.column = [self.name, self.supplier, str(self.poommok), str(self.quantity), str(self.multiply)]

    def __repr__(self):
        return f'{self.name} {self.supplier} {self.poommok} {self.quantity} {self.barcode}'




class ShoppingCart:

    def __init__(self):
        self.shoppingCart = []
        self.total = 0
        self.text = ""                      #Register 버튼 누를 시 나오는 텍스트
        self.Timestamp = datetime.now()

    def __repr__(self):
        return f'{self.shoppingCart} {self.total} {self.text} {self.Timestamp}'

    def find_item(self, bc):
        demo_list = []
        for product in product_list:
            if product.barcode == bc:
                demo_list.append(product)
        if len(demo_list) == 1:
            return demo_list[0]
        elif len(demo_list) > 1 :
            return demo_list

    def add_item(self, Product):
        self.shoppingCart.append(Product)
        self.total = 0
        for i in self.shoppingCart :
            self.total += i.multiply 

    def delete_shopping_cart(self, order):  #삭제버튼 1~5를 받는 함수. order를 통해 몇번째 항목의 삭제버튼이 클릭되었는지 받아온다.
        try:
            del self.shoppingCart[order]
            self.total = 0
            for i in self.shoppingCart :
                self.total += i.multiply 
        except IndexError:                  #물품이 차있지 않은 삭제버튼을 클릭하면 쇼핑카트에 들어있는 인덱스값보다 큰 값을 지우고자 하므로 IndexError가 작동한다.
            pass                            #이를 방지하기 위해 try, except 사용

    def clear_shopping_cart(self):
        self.shoppingCart.clear()
        self.total = 0
        
    def checkout(self, ssn):
        if self.total < 5:
            self.text = "You need more items"
        elif self.total ==5:
            for SSN in sheet[ssn_col]:
                if str(SSN.value) == ssn:
                    row = str(SSN.row)
                    sheet[visit_date + row].value = self.Timestamp.date()
                    sheet[visit_time + row].value = self.Timestamp.strftime("%I:%M %p")
                    sheet[visit_date + row].alignment = Alignment(horizontal = "center", vertical = "center")
                    sheet[visit_time + row].alignment = Alignment(horizontal = "center", vertical = "center")
                    count = 0  
                    k = len(self.shoppingCart)
                    while count < k :  #장바구니 안에있는 물품의 수만큼 아래의 코드 반복, 미리 만들어둔 Nlist, Qlist의 인덱스값을 이용하여 반복문을 사용하고 엑셀의 기둥을 지정한다.
                        sheet[Nlist[count] + row].value = self.shoppingCart[count].name + "/" + self.shoppingCart[count].supplier   # Nlist = [p1_name, p2_name, p3_name, p4_name, p5_name]
                        sheet[Qlist[count] + row].value = int(self.shoppingCart[count].quantity)                # Qlist = [p1_quantity, p2_quantity, p3_quantity, p4_quantity, p5_quantity]
                        sheet[Nlist[count] + row].alignment = Alignment(horizontal = "center", vertical = "center")
                        sheet[Qlist[count] + row].alignment = Alignment(horizontal = "center", vertical = "center")
                        count += 1
                    self.clear_shopping_cart()
                    break
            wb.save("test.xlsx")
            lookup_dict[ssn][1] = self.Timestamp.strftime("%Y-%m-%d")  # 코드 내에서 사용하는 딕셔너리에도 새롭게 등록된 이용자의 이용일자 추가. 코드 작동중 재입력 방지
            self.text = "Perfect"
        else:
            self.text = "Too many items"


class FirstWindow(QWidget, FirstWindow_formclass):  #이름 검색창 formclass = 디자인해둔 윈도우

    def __init__(self):
        super().__init__()
        self.setupUi(self)                          #디자인해둔  UI 셋업
        self.setWindowTitle("SSN SEARCH")
        self.connectFunction()
        self.x_pos = 650
        self.y_pos = 100

    def connectFunction(self):                      #특정동작(엔터키나 더블클릭 등)시 작동할 함수 연결
        self.nameInput.returnPressed.connect(self.ssnPrint)   
        self.ssnList.itemDoubleClicked.connect(self.BuildSecondWindow)
        self.ssnList.clicked.connect(self.visitDatePrint)

    def ssnPrint(self):
        self.ssnList.clear()
        self.info_clear()
        name = self.nameInput.text()
        for ssn, infoList in lookup_dict.items():               #infoList = [이순자, visit date, delegate, expiry, phone number]
            if infoList[0] == name and infoList[1] == None:     #이름으로 검색, 사용안한 이용자, infolist[1] = visit date 엑셀에 방문일자 기입여부에 따라서 이번달 사용인원 파악
                self.ssnList.addItem(ssn)
            elif infoList[0] == name and infoList[1]:           #이름으로 검색, 사용한 이용자
                self.ssnList.insertItem(0,ssn)  
                self.ssnList.item(0).setBackground(QColor('#FA8072')) 
            elif infoList[2] == name and infoList[1] == None:   #대리인으로 검색, 사용안한 이용자
                self.ssnList.addItem(ssn)
            elif infoList[2] == name and infoList[1]:           #대리인으로 검색, 사용한 이용자
                self.ssnList.insertItem(0,ssn)
                self.ssnList.item(0).setBackground(QColor('#FA8072'))
            elif infoList[4] == name and infoList[1] == None:   #전화번호로 검색, 사용안한 이용자
                self.ssnList.addItem(ssn)
            elif infoList[4] == name and infoList[1]:           #전화번호로 검색, 사용한 이용자
                self.ssnList.insertItem(0,ssn)
                self.ssnList.item(0).setBackground(QColor('#FA8072'))
            else:
                pass      
        if self.ssnList.item(0):                    #0번쨰 줄에 item이 없으면(lookup_dict에 이름이 없으면) "No value found" 출력
            pass
        else:
            self.ssnList.addItem("No value found")
    
    def visitDatePrint(self):
        if self.ssnList.currentItem().text() == "No value found":
            pass
        else:
            self.infoTable.setItem(0, 0, QTableWidgetItem(lookup_dict[self.ssnList.currentItem().text()][0]))
            self.infoTable.setItem(0, 1, QTableWidgetItem(str(lookup_dict[self.ssnList.currentItem().text()][1])[:10]))
            self.infoTable.setItem(0, 2, QTableWidgetItem(lookup_dict[self.ssnList.currentItem().text()][2]))
            self.infoTable.setItem(0, 3, QTableWidgetItem(str(lookup_dict[self.ssnList.currentItem().text()][3])))
            self.infoTable.setItem(0, 4, QTableWidgetItem(lookup_dict[self.ssnList.currentItem().text()][4]))

    def info_clear(self):
        self.infoTable.clearContents()

    @pyqtSlot(QListWidgetItem)
    def BuildSecondWindow(self, ssn):
        if self.ssnList.currentItem().text() == "No value found" or lookup_dict[self.ssnList.currentItem().text()][1]:
            pass
        else:
            secondWindow = SecondWindow(ssn.text(), self)
            secondWindow.move(self.x_pos, self.y_pos)
            secondWindow.show()
            if self.y_pos == 100:
                self.y_pos += 370
            elif self.x_pos == 650 and self.y_pos == 470:
                self.x_pos += 630
                self.y_pos = 100
            else:
                self.x_pos = 650
                self.y_pos = 100
            self.ssnList.clear()
            self.info_clear()
        

class SecondWindow(QMainWindow, SecondWindow_formclass): #바코드 입력창

    def __init__(self, ssn, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.ssn = ssn
        self.setLabel()
        self.connectFunction()
        self.cart = ShoppingCart()
        self.barcodeInput.setFocus()

    def setLabel(self):
        self.ssnLabel.setText(self.ssn)                    #받아온 SSN으로 주민번호라벨 표시
        self.nameLabel.setText(lookup_dict[self.ssn][0])   #받아온 SSN으로 이름라벨 표시
        self.delegate_label.setText(lookup_dict[self.ssn][2])
        self.expiry_label.setText(str(lookup_dict[self.ssn][3]))
        self.phone_label.setText(lookup_dict[self.ssn][4])
        #barcodeTable 헤더 사이즈 조정
        header = self.barcodeTable.horizontalHeader()       
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.ResizeToContents)


    def connectFunction(self): 
        self.barcodeInput.returnPressed.connect(self.enterInfo)
        self.registerButton.clicked.connect(self.register)
        self.clearButton.clicked.connect(self.clearing)
        self.delete_1.clicked.connect(self.delete_first)
        self.delete_2.clicked.connect(self.delete_second)
        self.delete_3.clicked.connect(self.delete_third)
        self.delete_4.clicked.connect(self.delete_fourth)
        self.delete_5.clicked.connect(self.delete_fifth)
    
    def enterInfo(self):
        self.registerLabel.setText("Press to register")
        self.barcode_label.clear()
        barcode = self.barcodeInput.text()
        try:
            if len(self.cart.shoppingCart) >4:                   #shoppingCart_dict에 5개이상 담겼을 경우 더이상 실행되지 않음
                self.barcodeInput.clear()
            else:
                if type(self.cart.find_item(barcode)) == list:
                    thirdWindow = ThirdWindow(self.cart.find_item(barcode), self)
                    chosed_index = thirdWindow.index
                    thirdWindow.show()
                    while chosed_index == None:
                        pass                   
                    self.cart.add_item(self.cart.find_item(barcode)[chosed_index])
                else:
                    self.cart.add_item(self.cart.find_item(barcode))
                self.show_cart()
        except KeyError:                                         #등록되지 않은 바코드가 입력될 경우 KeyError 발생. 이 경우 라벨에 "No barcode exist"출력
            self.barcode_label.setText("No barcode exist")
            self.barcodeInput.clear()

    def show_cart(self):
        self.barcodeTable.clearContents()
        for row,product in enumerate(self.cart.shoppingCart):               
            for col, value in enumerate(product.column):               
                self.barcodeTable.setItem(row, col, QTableWidgetItem(value))
                self.barcodeInput.clear()
        self.lcd_set(self.cart.total)

    def lcd_set(self, total):
        self.totalNumber.display(total)
        if total < 5:
            self.totalNumber.setStyleSheet("""QLCDNumber { 
                                            background-color: white; 
                                            color: black; }""")
        elif total == 5:
            self.totalNumber.setStyleSheet("""QLCDNumber { 
                                            background-color: #40FF00; 
                                            color: black; }""")
        else:
            self.totalNumber.setStyleSheet("""QLCDNumber { 
                                            background-color: #FA8072; 
                                            color: black; }""")

    #전체삭제
    def clearing(self):
        self.cart.clear_shopping_cart() 
        self.show_cart()
        self.barcodeInput.setFocus()

    #등록
    def register(self):  
        self.cart.checkout(self.ssn)
        self.registerLabel.setText(self.cart.text)
        self.show_cart()
        self.barcodeInput.setFocus()
        if self.cart.text == "Perfect":
            self.close()

    #삭제버튼 1~5
    def delete_first(self):
        self.cart.delete_shopping_cart(0)
        self.show_cart()
        self.barcodeInput.setFocus()

    def delete_second(self):
        self.cart.delete_shopping_cart(1)
        self.show_cart()
        self.barcodeInput.setFocus()

    def delete_third(self):
        self.cart.delete_shopping_cart(2)
        self.show_cart()
        self.barcodeInput.setFocus()

    def delete_fourth(self):
        self.cart.delete_shopping_cart(3)
        self.show_cart()
        self.barcodeInput.setFocus()

    def delete_fifth(self):
        self.cart.delete_shopping_cart(4)
        self.show_cart()
        self.barcodeInput.setFocus()


class ThirdWindow(QDialog, ThirdWindow_formclass):

    def __init__(self, lst, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.lst = lst
        for p in self.lst:
            self.overlap_list.addItem(p.name + "/" + p.supplier)
        self.overlap_list.itemDoubleClicked.connect(self.send_info)
        self.index = None

    def send_info(self):
        self.index = self.ssnList.currentRow()
        self.close





#실행 코드
create_plist()
create_dict()
window()




#ARCHIVES

# def create_pdict():  
#     f =  open('products.csv') 
#     csv_f = csv.reader(f)        # [['동서 보리차', 'CA마트', 1, 1, '8801037001808'], ['3분 짜장', 'CA마트', 1, 2, '8801045291314']]
#     for i in csv_f:
#         product_list[i[4]] = Product(i)