import xlsxwriter    
import os
import datetime



if os.path.exists("todaysBank.xlsx"):
    os.remove("todaysBank.xlsx")

player = 'peanuts'
class Item:
    def __init__(self, quantity, name, url, holder):
        self.quantity = quantity
        self.name = name
        self.url = url
        self.holder = holder


    def __lt__(self, other):
        return self.name < other.name

    def __str__(self):
        return f"{self.quantity}x[{self.name}]@{self.url} held by {self.holder}"


class parseBank:
    def __init__(self):
        self.bank = {}
        self.book = xlsxwriter.Workbook('todaysBank.xlsx')     
        self.sheet = self.book._add_sheet(f"{player}'s bank")     


    def parse(self, textfile):
        holder = "unknown"
        with open(textfile, 'r') as my_file:
            for line in my_file:
                if line[0:2] == "**":
                    holder = line[2:].split(':')[0]
                if line[0] == "*" or line[0] == '\n':
                    pass
                else:
                    line = line[2:] # first two chars always junk
                    quantity, rest = line.split('[') # split the quantity off
                    quantity = quantity[:-1] # take the x after the number.
                    name, rest = rest.split(']') # split name from rest
                    url = rest[1:-3] # remove () from around the link

                    tempo = Item(quantity, name, url, holder)
                            
                    if tempo.name in self.bank:
                        self.bank[tempo.name].quantity = int(self.bank[tempo.name].quantity) + int(tempo.quantity)
                    else:
                        self.bank[tempo.name] = tempo




    def printBank(self):
        for item in self.bank.values():
            print(item)


    def writeBankToXlsx(self):
        row = 1
        self.sheet.write(0, 0, 'Quantity')
        self.sheet.write(0, 1, 'Name')
        self.sheet.write(0, 2, 'Url')
        self.sheet.write(0, 3, 'Holder')
        for item in sorted(self.bank.values()) :     
            # print(item.name)
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




bank = parseBank()
bank.parse('bank.txt')
# bank.printBank()
bank.writeBankToXlsx()
