import xlsxwriter    
import os
import datetime


player = 'peanuts' # change this






if os.path.exists("todaysBank.xlsx"):
    os.remove("todaysBank.xlsx")


class Item:
    def __init__(self, quantity, name, url, holder):
        self.quantity = quantity
        self.name = name
        self.url = url
        self.holder = holder


class parseBank:
    def __init__(self, player):
        self.bank = {}
        self.book = xlsxwriter.Workbook('todaysBank.xlsx')     
        self.sheet = self.book._add_sheet(f"{player}'s bank")     


    def parse(self, textfile):
        with open(textfile, 'r') as my_file:
            for line in my_file:
                if line[0] == "*" or line[0] == '\n':
                    pass
                else:
                    line = line[2:]
                    quantity, rest = line.split('[')
                    quantity = quantity[:-1]
                    name, rest = rest.split(']')
                    url = rest[1:-1]
                    # print(quantity, name, url)
                    a = Item(quantity, name, url, 'peanuts')
                            
                    if a.name in self.bank:
                        self.bank[a.name].quantity = int(self.bank[a.name].quantity) + int(a.quantity)
                    else:
                        self.bank[a.name] = a




    def printBank(self):
        for item in self.bank.values():
            print(f'q: {item.quantity}, name: {item.name}, url: {item.url}')


    def writeBankToXlsx(self):
        row = 1
        self.sheet.write(0, 0, 'Quantity')
        self.sheet.write(0, 1, 'Name')
        self.sheet.write(0, 2, 'Url')
        self.sheet.write(0, 3, 'Holder')
        for item in self.bank.values() :     
      
            # write operation perform     
            self.sheet.write(row, 0, item.quantity)
            self.sheet.write(row, 1, item.name)
            self.sheet.write(row, 2, item.url)
            self.sheet.write(row, 3, item.holder)     
              
            # incrementing the value of row by one with each iterations.     
            row += 1    
              
        time = datetime.datetime.now()

        self.sheet.write(row, 0, f'DDMMYYYY {time.day}-{time.month}-{time.year}')

        self.book.close()  




bank = parseBank(player)
bank.parse('bank.txt')
bank.printBank()
bank.writeBankToXlsx()
# ite = item()