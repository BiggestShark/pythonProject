import pandas as pd ## 用於讀取檔案
import matplotlib.pyplot as plt ## 用於製作圖表
import os ## 用於獲取檔案名稱
import sqlite3 as sq ## 資料庫
import tkinter as tk
from tkinter import filedialog , messagebox

def set_chinese_font():
    import matplotlib.font_manager as fm
    simhei_font = fm.FontProperties( fname = 'C:/Windows/Fonts/微軟正黑體.ttf' )
    plt.rcParams['font.family'] = simhei_font.get_name()

def line_chart( data ):
    set_chinese_font()
    for column in data.columns[1:]:
        plt.plot( data[data.columns[0]] , data[column] , marker = "o" , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( "折線圖" )
    plt.legend()
    plt.grid( True )
    plt.show()

def scatter_chart( data ):
    set_chinese_font()
    for column in data.columns[1:]:
        plt.scatter( data[data.columns[0]] , data[column] , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( '點狀圖' )
    plt.legend()
    plt.grid(True)
    plt.show()

def bar_chart( data ):
    set_chinese_font()
    for column in data.columns[1:]:
        plt.bar( data[data.columns[0]] , data[column] , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( "柱狀圖" )
    plt.legend()
    plt.grid(True)
    plt.show()

def histogram( data ):
    set_chinese_font()
    for column in data.columns[1:]:
        plt.hist( data[column] , bins=10 , alpha = 0.5 , label = column )
    plt.xlabel('Value')
    plt.ylabel('Frequency')
    plt.title('Histogram')
    plt.legend()
    plt.grid(True)
    plt.show()

def pie_chart( data ):
    set_chinese_font()
    # 使用第一列數據
    if len( data.columns ) > 1:
        plt.pie( data[data.columns[1]] , labels = data[data.columns[0]] , autopct = "%1.1f%%" )
        plt.title( "圓餅圖" )
        plt.show()
    else:
        print( "Data not sufficient for pie chart" )

def load_excel_file():
    filepath = filedialog.askopenfilename( filetypes = [( "Excel files" , "*.xlsx *.xls" )])
    if not filepath:
        return
    try:
        global excel_data
        excel_data = pd.read_excel(filepath, sheet_name=None)
        sheet_names = list(excel_data.keys())
        sheet_menu['menu'].delete(0, 'end')
        for sheet_name in sheet_names:
            sheet_menu['menu'].add_command(label=sheet_name, command=tk._setit(selected_sheet, sheet_name))
        selected_sheet.set(sheet_names[0])
        messagebox.showinfo("加载成功", f"成功加载Excel文件：{filepath}")
    except Exception as e:
        messagebox.showerror("加载失败", f"无法加载Excel文件：{str(e)}")

def plot_selected_chart():
    chart_type = selected_chart.get()
    sheet_name = selected_sheet.get()
    df = excel_data[sheet_name]
    if chart_type == "折线图":
        line_chart(df)
    elif chart_type == "点状图":
        scatter_chart(df)
    elif chart_type == "柱状图":
        bar_chart(df)
    elif chart_type == "直方图":
        histogram(df)
    elif chart_type == "圆饼图":
        pie_chart(df)
    else:
        messagebox.showerror( "错误", "无效的图表类型" )

## db_path = "C:/Users/User/Documents/Python Scripts/sqlite"
"""
print( "輸入文件路徑：" )
path = input()
file_name = os.path.basename( path ) ## 取得檔案名稱
data = pd.read_excel( path , sheet_name = None ) ## 取得所有工作表

print( "可用工作表：" )
sheet_names = list( data.keys() )
for i , sheet_name in enumerate( sheet_names ):
    print( f"{i + 1}. {sheet_name}" )

print( "選擇工作表( 輸入編號 )：" )
sheet_index = int( input().strip() ) - 1

df = data[sheet_names[sheet_index]]
print( "選擇的工作表數據：" )
print( df )

## db_name = os.path.splitext( file_name )[0] + '.db' ## 設定資料庫名稱
## conn = sq.connect( db_name ) ## 連接資料庫
## for sheet_name , df in data.items():
##    df.to_sql( sheet_name , conn , if_exists = 'replace' , index = False )

print( "\n選擇圖表類型( 輸入編號 )：" )
print( "1. 折線圖" )
print( "2. 點狀圖" )
print( "3. 柱狀圖" )
print( "4. 直方圖" )
print( "5. 圓餅圖" )

chart_type = int( input().strip() )

if chart_type == 1:
    line_chart( df )
elif chart_type == 2:
    scatter_chart( df )
elif chart_type == 3:
    bar_chart( df )
elif chart_type == 4:
    histogram( df )
elif chart_type == 5:
    pie_chart( df )
else:
    print( "編號錯誤" )
"""
# 创建主窗口
root = tk.Tk()
root.title("Excel 数据可视化")

# 创建并布局按钮和菜单
load_button = tk.Button(root, text="加载Excel文件", command=load_excel_file)
load_button.pack(pady=10)

sheet_label = tk.Label(root, text="选择工作表：")
sheet_label.pack()
selected_sheet = tk.StringVar(root)
sheet_menu = tk.OptionMenu(root, selected_sheet, "")
sheet_menu.pack(pady=5)

chart_label = tk.Label(root, text="选择图表类型：")
chart_label.pack()
chart_options = ["折线图", "点状图", "柱状图", "直方图", "圆饼图"]
selected_chart = tk.StringVar(root)
selected_chart.set(chart_options[0])
chart_menu = tk.OptionMenu(root, selected_chart, *chart_options)
chart_menu.pack(pady=5)

plot_button = tk.Button(root, text="绘制图表", command=plot_selected_chart)
plot_button.pack(pady=10)

# 运行主窗口
root.mainloop()

## conn.close()