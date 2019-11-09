import sys
import matplotlib as mpl
import matplotlib.pyplot as plt
from openpyxl import load_workbook
from datetime import datetime
from functools import reduce


######  TenlongOrdersPlot Class  ######

class TenlongOrdersPlot:
    def __init__(self, xlsx_filename):
        ws = self.loadXLSX(xlsx_filename)
        self.orders = self.getOrders(ws)

    def loadXLSX(self, xlsx_filename):
        wb = load_workbook(xlsx_filename)
        ws = wb.active
        return ws

    def getOrders(self, ws):
        orders = []
        idx = 0
        for row in ws.rows:
            if idx != 0:
                dateStr = row[-1].value
                subtotal = row[-3].value
                datetimeObj = datetime.strptime(dateStr, '%Y/%m/%d %I:%M %p')
                orders.append( (datetimeObj, subtotal) )
            idx = idx + 1
        return orders
    
    def costByYear(self):
        print("Compute cost by year...")
        dict = {}
        for row in self.orders:
            year = row[0].year
            subtotal = row[1]
            if year in dict:
                dict[year] = dict[year] + subtotal
            else:
                dict[year] = subtotal
        return dict

    def drawPlotBar(self, x, y, labels, x_title, y_title, top_title, average=0):
        fig, ax = plt.subplots()
        plt.title(top_title)
        plt.xlabel(x_title)
        plt.ylabel(y_title)
        plt.grid(True, axis="y", ls=":", color="r", alpha=0.35)
        bar_plot = plt.bar(x, y, align="center", color="b", tick_label=labels, alpha=0.8)

        for idx,rect in enumerate(bar_plot):
            height = rect.get_height()
            ax.text(rect.get_x() + rect.get_width()/2., 1.0*height,
                    y[idx],
                    ha='center', va='bottom', rotation=0)
        if average != 0:
            average_label = "average: " + str( int(average) )
            plt.axhline(y=average, c="g", ls="-.", lw=1, alpha=0.9, label=average_label)
            labels = ["average"]
            handles, _ = ax.get_legend_handles_labels()
            plt.legend(handles = handles[1:], labels = labels)
        plt.show()

    def yearListWithRange(self, start_year, end_year):
        if end_year > start_year:
            year_range = range(start_year, end_year+1)
        else:
            year_range = range(start_year, end_year-1, -1)
        return list(year_range)
    
    def showYearCostBar(self, years):
        x = []
        y = []
        labels = []
        idx = 1
        for year in years:
            labels.append(year)
            x.append(idx)
            if year in self.year_cost_dict:
                y.append(self.year_cost_dict[year])
            else:
                y.append(0)
            idx += 1
        self.drawPlotBar(x, y, labels, "Year", "NT$", "Cost by year")

    def showBar(self, start_year = None, end_year = None):
        if not hasattr(self, 'year_cost_dict'):
            self.year_cost_dict = self.costByYear()
        if len(self.year_cost_dict) == 0:
            print("ERROR: 無可用資料！")
            return
        print(self.year_cost_dict)

        if start_year != None and end_year != None:
            self.showYearCostBar( self.yearListWithRange(start_year, end_year) )
        else:
            sort_yaers = sorted(list(self.year_cost_dict.keys()))
            self.showYearCostBar( self.yearListWithRange(sort_yaers[0], sort_yaers[-1]) )

    def showBarWithDiscrete(self, years):
        if not hasattr(self, 'year_cost_dict'):
            self.year_cost_dict = self.costByYear()
        if len(self.year_cost_dict) == 0:
            print("ERROR: 無可用資料！")
            return
        print(self.year_cost_dict)
        self.showYearCostBar(years)

    def monthCostInYear(self, year):
        print("Compute cost in the " + str(year) + "...")
        dict = {}
        for m in range(1,12+1):
            dict[m] = 0
        for row in self.orders:
            if row[0].year == year:
                m = row[0].month
                subtotal = row[1]
                dict[m] = dict[m] + subtotal
        return dict

    def showMonthCostBar(self, year):
        month_cost_dict = self.monthCostInYear(year)
        print(month_cost_dict)
        x = []
        y = []
        for m in range(1,12+1):
            x.append(m)
            y.append(month_cost_dict[m])
        total = reduce((lambda a, b: a + b), list(month_cost_dict.values()))
        title = "Cost in the " + str(year) + " C.E.  (total: NT$ " + str(total) + ")"
        average = total / 12 
        self.drawPlotBar(x, y, x, "Month", "NT$", title, average)


#####  TenlongOrdersPlot Class End  #####


def handleParameters():
    if len(sys.argv) == 1:
        return (True, "提示：請指定檔名。", "", [])
    elif len(sys.argv) == 2:
        xlsx_filename = sys.argv[1]
        return (False, "", xlsx_filename, [])
    elif len(sys.argv) == 3:
        xlsx_filename = ""
        param_str = ""
        param_list = []

        for i in range(1,2+1):
            if ".xlsx" in sys.argv[i]:
                xlsx_filename = sys.argv[i]
            if ".xlsx" not in sys.argv[i]:
                param_str = sys.argv[i]

        if '-' in param_str:
            param_list.append('-')
            split_list = param_str.split('-')
            for s in split_list:
                if s.strip() != '':
                    param_list.append( int(s.strip()) )
        elif ',' in param_str:
            param_list.append(',') 
            split_list = param_str.split(',')
            for s in split_list:
                if s.strip() != '':
                    param_list.append( int(s.strip()) )
        elif param_str.isdigit():
            param_list.append('m')
            param_list.append( int(param_str) )
        return (False, "", xlsx_filename, param_list)
    else:
        return (True, "ERROR: 錯誤的參數！", "", [])


def main():
    is_err, msg, xlsx_filename, show_years = handleParameters()
    if is_err:
        print(msg)
        return
    print()
    orders_plot = TenlongOrdersPlot(xlsx_filename)

    if len(show_years) >= 2:
        head, tail = show_years[0], show_years[1:]
        if head == '-':
            start_year = tail[0]
            if len(tail) > 1:
                end_year = tail[1]
            else:
                end_year = datetime.now().year
            orders_plot.showBar(start_year, end_year)
        elif head == ',':
            orders_plot.showBarWithDiscrete(tail)
        elif head == 'm':
            orders_plot.showMonthCostBar(tail[0])
    else:
        orders_plot.showBar()   # show all years
    print()


# Run
main()
