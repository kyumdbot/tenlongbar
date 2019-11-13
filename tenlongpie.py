import sys
import json
import matplotlib as mpl
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from datetime import datetime
from functools import reduce


######  TenlongOrdersPlotPie Class  ######

class TenlongOrdersPlotPie:
    def __init__(self, xlsx_filename, category_data):
        ws = self.loadXLSX(xlsx_filename)
        self.orders = self.getOrders(ws)
        self.category_data = category_data 

    def loadXLSX(self, xlsx_filename):
        wb = load_workbook(xlsx_filename)
        ws = wb.active
        return ws

    def getOrders(self, ws):
        orders = []
        idx = 0
        for row in ws.rows:
            if idx != 0:
                date_str = row[-1].value
                subtotal = row[-3].value
                if type(row[-7]) is str:
                    publisher = row[-7].value.strip()
                elif row[-7] is not None:
                    publisher = str(row[-7].value)
                else:
                    publisher = ""
                datetime_obj = datetime.strptime(date_str, '%Y/%m/%d %I:%M %p')
                orders.append( (datetime_obj, subtotal, publisher) )
            idx = idx + 1
        return orders

    def show(self, year):
        sub_orders = []
        for row in self.orders:
            if row[0].year == year:
                sub_orders.append(row)
        if len(sub_orders) == 0:
            print("此年份無資料: " + str(year))
            return

        category_names = list(self.category_data['category'].keys())
        total = 0
        subtotal_dict = {}
        for name in category_names:
            subtotal_dict[name] = 0
        
        # row => (datetime, subtotal, publisher)
        for row in sub_orders:
            total = total + row[1]
            for name in category_names:
                if row[2] in self.category_data['category'][name]:
                    subtotal_dict[name] += row[1]
                    break
        print(subtotal_dict)
        
        categories_subtotal = reduce((lambda x, y: x + y), list(subtotal_dict.values()))
        default_subtotal = total - categories_subtotal

        labels = [ self.category_data["default"] ]
        labels.extend(category_names)

        subtotal_labels = [ default_subtotal ]
        default_rate = default_subtotal / total 
        rate_list = [ default_rate ]
        for name in category_names:
            subtotal_labels.append(subtotal_dict[name])
            rate_list.append( subtotal_dict[name]/total )
        print(rate_list)
        
        subtotal_labels = [ "NT$ " + str(item) for item in subtotal_labels]

        plt.pie(rate_list, labels=subtotal_labels, autopct="%3.2f%%",
                startangle=45, pctdistance=0.7, labeldistance=1.2)
        top_title = "Cost in the " + str(year) + " C.E.  (total: NT$ " + str(total) + ")"
        plt.title(top_title)

        plt.legend(labels=labels, loc="upper center", bbox_to_anchor=(1,0,0.1,1))
        plt.show()
                
#####  TenlongOrdersPlotPie Class End  #####



def handleParameters():
    if len(sys.argv) == 1 or len(sys.argv) == 2:
        return (True, "提示：請指定檔名與年份。", "", 0)
    elif len(sys.argv) == 3:
        xlsx_filename = ""
        param_str = ""

        for i in range(1,2+1):
            if ".xlsx" in sys.argv[i]:
                xlsx_filename = sys.argv[i]
            if ".xlsx" not in sys.argv[i]:
                param_str = sys.argv[i]
        
        if param_str.isdigit():
            return (False, "", xlsx_filename, int(param_str))
        else:
            return (True, "ERROR: 錯誤的參數！", xlsx_filename, 0)  
    else:
        return (True, "ERROR: 錯誤的參數！", "", 0)


def main():
    is_err, msg, xlsx_filename, year = handleParameters()
    if is_err:
        print(msg)
        return
    print()

    with open('category.json', encoding='utf-8') as f:
        category_data = json.load(f)
    print(category_data)
    print()

    plot_pie = TenlongOrdersPlotPie(xlsx_filename, category_data)

    if year != 0:
        plot_pie.show(year)    
    print()


# Run
main()

