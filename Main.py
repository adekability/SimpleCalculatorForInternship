import tkinter as tk #библиотека графического интерфейса
import re#бибилиотека регулярных выражений
import math
import xlrd
import xlwt
from tkinter.filedialog import askopenfile
class main:#основной класс калькулятора
    def make(self):
        filename = askopenfile()
        filename=filename.name
        rb = xlrd.open_workbook(filename)
        sheet = rb.sheet_by_index(0)
        sums = []
        for rownum in range(sheet.nrows):
            sums.append([])
            row = sheet.row_values(rownum)
            for col in row:
                sums[rownum].append(col)
                print(col,end="\t")
            print()
        oper = ["+","-","*","/"]
        for y in range(len(sums)):
            endup = 0
            for i in range(len(sums[y])-1):
                if i == 0:
                    if oper[y] == "+":
                        endup += sums[y][i] + sums[y][i + 1]
                    elif oper[y] == "-":
                        endup += sums[y][i] - sums[y][i + 1]
                    elif oper[y] == "*":
                        endup += sums[y][i] * sums[y][i + 1]
                    elif oper[y] == "/":
                        endup += sums[y][i] / sums[y][i + 1]
                else:
                    if oper[y] == "+":
                        endup += sums[y][i + 1]
                    elif oper[y] == "-":
                        endup -= sums[y][i + 1]
                    elif oper[y] == "*":
                        endup *= sums[y][i + 1]
                    elif oper[y] == "/":
                        endup /= sums[y][i + 1]

            sums[y].append(endup)
        for i in range(len(sums)-1):
            for y in range(len(sums[i])):
                if sums[i][y]-round(sums[i][y])==0:
                    sums[i][y]=int(sums[i][y])
        print(sums)
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Лист1")
        for i in range(len(sums)):
            for y in range(len(sums[i])):

                word = str(sums[i][y])

                ws.write(i,y,word)
        wb.save("file.xls")
    def __init__(self):#метод init, в котором будет окно приложения
        self.root = tk.Tk()#вызов метода Tk() для создания окна
        self.root.geometry("300x510")#разрешение окна 300 по ширине, и 510 по высоте
        self.root.title("Калькулятор")#заголовок окна
        self.mainMenu = tk.Menu(self.root)
        self.mainMenu.add_command(label="Добавить файл",command=self.make)
        self.root.config(menu=self.mainMenu)



        #создаем поле ввода чисел и операций
        self.entry = tk.Entry(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                              font=("Helvetica",34),#шрифт Helveticа, размер 34
                              justify="right")#выравнивание ввода вправо
        self.entry.pack(side="top")#расположение виджета ввода сверху окна

        # создаем кнопку взятия процента числа
        self.percentButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="%",#надпись % в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type("%",percentage = True))#команда вызова функции ввода %
        self.ceButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки#рельеф кнопки
                                       width=4,#ширина кнопки
                                       text="CE",#надпись CE в кнопке
                                       font=("Helvetica",20))#шрифт Helveticа, размер 20
        self.cButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки#рельеф кнопки
                                       width=4,#ширина кнопки
                                       text="C",#надпись C в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=self.clear)#команда вызова функции очистки ввода
        self.deleteButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="←",#надпись ← в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=self.delete)#команда вызова функции удаления
        self.oneOnXButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="1/x",#надпись 1/x в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type(str(round(1/float(self.entry.get()),5)),operation=True))
        self.xPow2Button = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="x^2",#надпись x^2 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type(str(math.pow(float(self.entry.get()),2)),operation=True))
        self.xSqrtButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="2√x",#надпись 2√x в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type(str(round(math.sqrt(int(self.entry.get())),5)), operation=True))
        self.divButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="/",#надпись / в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type("/"))#команда вызова функции деления
        self.sevenButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="7",#надпись 7 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('7'))#команда вызова функции ввода 7
        self.eightButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="8",#надпись 8 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('8'))#команда вызова функции ввода 8
        self.nineButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="9",#надпись 9 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('9'))#команда вызова функции ввода 9
        self.multiplyButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="*",#надпись * в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type("*"))#команда вызова функции вычисления умножения
        self.fourButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="4",#надпись 4 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('4'))#команда вызова функции ввода 4
        self.fiveButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="5",#надпись 5 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('5'))#команда вызова функции ввода 5
        self.sixButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="6",#надпись 6 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('6'))#команда вызова функции ввода 6
        self.minusButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="-",#надпись - в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type("-"))#команда вызова функции вычитания
        self.oneButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="1",#надпись 1 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('1'))
        self.twoButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="2",#надпись 2 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('2'))#команда вызова функции ввода 2
        self.threeButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="3",#надпись 3 в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('3'))#команда вызова функции ввода 3
        self.plusButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="+",#надпись + в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type("+"))#команда вызова функции вычисления +
        self.negPosButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="+ -",#надпись + - в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                      command=self.negative)
        self.nullButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="0",#надпись '0' в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=lambda: self.type('0'))#команда вызова функции ввода
        self.commaButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text=",",#надпись ',' в кнопке
                                       font=("Helvetica",20))#шрифт Helveticа, размер 20
        self.equalButton = tk.Button(self.root,#аргумент в виде переменной окна - виджет создается в этом окне
                                       relief="groove", #рельев кнопки
                                       width=4,#ширина кнопки
                                       text="=",#надпись = в кнопке
                                       font=("Helvetica",20),#шрифт Helveticа, размер 20
                                       command=self.equal)#команда вызова функции вычисления
        self.entry.insert(0,"0")#ввода цифры 0 в начало строки справа по умолчанию



        self.percentButton.place(x=1, y=180)#расположение кнопки % в окне
        self.ceButton.place(x=76, y=180)#расположение кнопки в окне
        self.cButton.place(x=151, y=180)#расположение кнопки очистки вводав окне
        self.deleteButton.place(x=226, y=180)#расположение кнопки удаления последнего ввода в окне

        self.oneOnXButton.place(x=1, y=235)#расположение кнопки 1/x в окне
        self.xPow2Button.place(x=76, y=235)#расположение кнопки в окне
        self.xSqrtButton.place(x=151, y=235)#расположение кнопки взятия из под корня в окне
        self.divButton.place(x=226, y=235)#расположение кнопки / в окне

        self.sevenButton.place(x=1, y=290)#расположение кнопки 7 в окне
        self.eightButton.place(x=76, y=290)#расположение кнопки 8 в окне
        self.nineButton.place(x=151, y=290)#расположение кнопки 9 в окне
        self.multiplyButton.place(x=226, y=290)#расположение кнопки * в окне

        self.fourButton.place(x=1, y=345)#расположение кнопки 4 в окне
        self.fiveButton.place(x=76, y=345)#расположение кнопки 5 в окне
        self.sixButton.place(x=151, y=345)#расположение кнопки 6 в окне
        self.minusButton.place(x=226, y=345)#расположение кнопки - в окне

        self.oneButton.place(x=1, y=400)#расположение кнопки 1 в окне
        self.twoButton.place(x=76, y=400)#расположение кнопки 2 в окне
        self.threeButton.place(x=151, y=400)#расположение кнопки 3 в окне
        self.plusButton.place(x=226, y=400)#расположение кнопки + в окне

        self.negPosButton.place(x=1, y=455)#расположение кнопки +- в окне
        self.nullButton.place(x=76, y=455)#расположение кнопки 0 в окне
        self.commaButton.place(x=151, y=455)#расположение кнопки ',' в окне
        self.equalButton.place(x=226, y=455)#расположение кнопки '=' в окне


        self.root.mainloop()#запуск цикла отображения окна

    def percentCount(self, string):
        add_string = string

        string.replace("%","")
        sum = ""
        for i in range(len(string)-1,-1,-1):
            if string[i]!="*" and string[i]!="/" and string[i]!="+" and string[i]!="-":
                sum+=string[i]
            else:
                break
        sum = [x for x in sum]
        sum.reverse()
        sum = ''.join(sum)
        add_string = add_string.replace(sum,"")
        sum = sum.replace("%","")
        sum = int(sum)/100
        add_string+=str(sum)
        return add_string
    def type(self,symbol, operation=None, percentage=None):#метод границ для чисел больше 10 знаков
        if type(symbol) is str:
            if symbol[len(symbol)-2:]==".0":
                symbol = float(symbol)
                symbol = int(symbol)
        if operation is True:
            self.entry.delete(0,"end")
        if percentage is True and symbol=="%":
            sum = self.percentCount(self.entry.get()+symbol)
            self.entry.delete(0, "end")
            symbol=sum

        if len(self.entry.get())>10:#если больше 10 знаков
            print("Нельзя набрать больше 10-значного числа!")#вывод
            return
        #else:#если меньше 10 знаков
        if self.entry.get()=="0":#если имеется ноль по умолчанию
            self.entry.delete(0)#удаление стоящего по умолчанию нуля
        self.entry.insert(len(self.entry.get()),str(symbol))#внести аргумент symbol в окно ввода


    def clear(self):#метод очистки ввода
        self.entry.delete(0,"end")#удалить от индекса 0 до конца
        self.entry.insert(len(self.entry.get())+1,0)#вставить 0 по умолчанию после очистки

    def delete(self):#метод удаления крайнего символа
        if len(self.entry.get()) == 1:#если до/после удаления остался 1 знак
            self.entry.delete(len(self.entry.get()) - 1, "end")#удалить крайний символ
            self.entry.insert(0,"0")#вставить ноль по умолчанию
        else:
            self.entry.delete(len(self.entry.get()) - 1, "end")#если нет, удалить крайний символ

    def negative(self):#метод изменения знака
        sum = int(self.entry.get())
        sum = -sum
        self.entry.delete(0,"end")
        self.entry.insert(len(self.entry.get())-1,sum)


    def equal(self):#метод вычисления операции
        countString = self.entry.get()#берем строку из ввода
        self.entry.delete(0,"end")#очищаем ввод
        if countString[0]=="-":#если первое число отрицательное - не допустить к списку знаков
            numList = re.findall(r"[\w']+", countString)  # убираем знаки и оставляем числа
            print(numList)
            numList = [int(y) for y in numList]  # переинициализация списка чисел
            numList[0] = -numList[0]#присваиваем этот знак первому числу из списка чисел
            opr = re.sub("[0-9]", "", countString)  # оставляем лишь знаки вычисления без чисел
            opr = [x for x in opr]  # переинициализация списка знаков вычисления
            opr.pop(0)#удаляем первый знак -
            sum = 0  # переменная для присваивания выходящего числа
        else:
            numList = re.findall(r"[\w']+", countString)#убираем знаки и оставляем числа
            print(numList)
            numList = [int(y) for y in numList]#переинициализация списка чисел
            opr = re.sub("[0-9]", "", countString)#оставляем лишь знаки вычисления без чисел
            opr = [x for x in opr]#переинициализация списка знаков вычисления
            sum = 0#переменная для присваивания выходящего числа

        for i in range(len(opr)):#цикл для прохождения по списку знаков
            if i==0:#если цикл только начался, присваиваем переменной sum операцию с первыми двумя элементами массива чисел
                if opr[i]=="+":#если знак +
                    sum+=numList[i]+numList[i+1]#присваиваем сложение
                elif opr[i]=="-":#если знак -
                    sum+=numList[i]-numList[i+1]#присваиваем вычитание
                elif opr[i]=="*":#если знак *
                    sum+=numList[i]*numList[i+1]#присваиваем умножение
                elif opr[i]=="/":#если знак /
                    sum+=numList[i]/numList[i+1]#присваиваем деление
                sum = round(sum, 5)
            else:#если цикл продолжается присваиваем операцию со следующим элементом напрямую
                if opr[i]=="+":#если знак +
                    sum+=numList[i+1]##присваиваем сложение
                elif opr[i]=="-":#если знак -
                    sum-=numList[i+1]#присваиваем вычитание
                elif opr[i]=="*":#если знак *
                    sum*=numList[i+1]#присваиваем умножение
                elif opr[i]=="/":#если знак /
                    sum/=numList[i+1]#присваиваем деление
                sum = round(sum, 5)
        self.entry.insert(len(self.entry.get())+1,str(sum))#вывод суммы во ввод
main()#вызов основного класса