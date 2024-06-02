import pandas as pd
import matplotlib.pyplot as plt
import os
import sqlite3 as sq
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def set_chinese_font():
    import matplotlib.font_manager as fm
    font_paths = [
        "C:/Windows/Fonts/msjh.ttc",
        "C:/Windows/Fonts/msjh.ttf",
        "C:/Windows/Fonts/msjhl.ttc",
        "C:/Windows/Fonts/msjhl.ttf",
        "C:/Windows/Fonts/微軟正黑體.ttf"
    ]
    for font_path in font_paths:
        if os.path.exists( font_path ):
            msjh_font = fm.FontProperties( fname = font_path )
            plt.rcParams["font.family"] = msjh_font.get_name()
            return
    raise FileNotFoundError( "無法找到'微軟正黑體'字體" )

def plot_chart( df , chart_type , x_column , y_column , x_interval , y_interval ):
    set_chinese_font()
    fig , ax = plt.subplots()

    if chart_type == "折線圖":
        ax.plot( df[x_column] , df[y_column] , marker = "o" , label = y_column )
        ax.set_title( "折線圖" )
    elif chart_type == "點狀圖":
        ax.scatter( df[x_column] , df[y_column] , label = y_column )
        ax.set_title( "點狀圖" )
    elif chart_type == "柱狀圖":
        ax.bar( df[x_column] , df[y_column] , label = y_column )
        ax.set_title( "柱狀圖" )
    elif chart_type == "圓餅圖":
        pie_columns_window = tk.Toplevel( root )
        pie_columns_window.title( "選擇圓餅圖資料行" )

        tk.Label( pie_columns_window , text = "選擇變數行:" ).grid( row = 0 , column = 0 , padx = 5 , pady = 5 )
        tk.Label( pie_columns_window , text = "選擇數據行:" ).grid( row = 1 , column = 0 , padx = 5 , pady = 5 )

        label_column = tk.StringVar( pie_columns_window )
        value_column = tk.StringVar( pie_columns_window )
        
        label_menu = tk.OptionMenu( pie_columns_window , label_column , *df.columns )
        value_menu = tk.OptionMenu( pie_columns_window , value_column , *df.columns )
        
        label_menu.grid( row = 0 , column = 1 , padx = 5 , pady = 5 )
        value_menu.grid( row = 1 , column = 1 , padx = 5 , pady = 5 )

        def plot_pie_chart():
            try:
                pie_labels = df[label_column.get()].values
                pie_values = df[value_column.get()].values
                ax.clear()
                ax.pie( pie_values , labels = pie_labels , autopct = "%1.1f%%" )
                ax.set_title( "圓餅圖" )
                canvas.draw()
                pie_columns_window.destroy()
            except KeyError as e:
                messagebox.showerror( "錯誤" , f"選擇的資料有誤: {e}" )
            except Exception as e:
                messagebox.showerror( "錯誤" , f"圓餅圖無法繪製: {e}" )

        tk.Button( pie_columns_window , text = "繪製圓餅圖" , command = plot_pie_chart ).grid( row = 2 , column = 0 , columnspan = 2 , pady = 10 )

    ax.legend()
    ax.grid( True )

    if x_interval > 0:
        ax.xaxis.set_major_locator( plt.MultipleLocator( x_interval ) )
    if y_interval > 0:
        ax.yaxis.set_major_locator( plt.MultipleLocator( y_interval ) )

    return fig

def load_excel_file():
    filepath = filedialog.askopenfilename( filetypes = [( "Excel files" , "*.xlsx *.xls" )] )
    if not filepath:
        return
    try:
        global data
        global file_name
        global sheet_names

        file_name = os.path.basename( filepath )

        # 先讀取Excel檔案的前幾列，讓使用者選擇資料從第幾列開始
        temp_data = pd.read_excel( filepath , sheet_name = None , nrows = 10 )
        
        sheet_names = list( temp_data.keys() )
        header_row = simpledialog.askinteger( "資料從第幾列開始？" , "請輸入數字(從1開始)：" , minvalue= 1 , maxvalue = 10 )
        if not header_row:
            return
        
        # 讀取所有檔案中的工作表
        data = pd.read_excel( filepath , sheet_name = None , header = header_row-1 )
        for sheet in data:
            # 刪除空白行、列
            data[sheet].dropna( how = "all" , inplace = True )
            data[sheet].dropna( axis=1 , how = "all" , inplace = True )

        for sheet in data:
            data[sheet] = data[sheet].dropna( how = "all" )

        sheet_names = list( data.keys() )

        sheet_menu["menu"].delete( 0 , "end" )
        for sheet_name in sheet_names:
            sheet_menu["menu"].add_command( label = sheet_name , command = tk._setit( selected_sheet , sheet_name , update_column_menus ) )
        selected_sheet.set( sheet_names[0] )

        update_column_menus()

        messagebox.showinfo( "載入成功" , f"成功載入Excel檔案：{filepath}" )
    except Exception as e:
        messagebox.showerror( "載入失敗" , f"無法載入Excel檔案：{str(e)}" )

def calculate_and_show_correlation( df , x_column , y_column ):
    correlation = df[x_column].corr( df[y_column] )
    messagebox.showinfo( "相關係數", f"{x_column} 和 {y_column} 之間的相關係數是: {correlation:.2f}" )

def update_column_menus( *args ):
    sheet_name = selected_sheet.get()
    if sheet_name:
        columns = data[sheet_name].columns
        x_column_menu["menu"].delete( 0 , "end" )
        y_column_menu["menu"].delete( 0 , "end" )
        for column in columns:
            x_column_menu["menu"].add_command( label = column , command=tk._setit( selected_x_column , column ) )
            y_column_menu["menu"].add_command( label = column , command=tk._setit( selected_y_column , column ) )
        selected_x_column.set( columns[0] )
        selected_y_column.set( columns[1] if len( columns ) > 1 else columns[0] )

def plot_selected_chart():
    chart_type = selected_chart.get()
    sheet_name = selected_sheet.get()
    x_column = selected_x_column.get()
    y_column = selected_y_column.get()
    x_interval = float( x_interval_entry.get() ) if x_interval_entry.get() else 1
    y_interval = float( y_interval_entry.get() ) if y_interval_entry.get() else 1

    df = data[sheet_name]

    if chart_type == "相關係數":
        calculate_and_show_correlation( df , x_column , y_column )
        return

    fig = plot_chart( df , chart_type , x_column , y_column , x_interval , y_interval )
    if fig:
        global canvas
        if canvas:
            canvas.get_tk_widget().grid_forget()
        canvas = FigureCanvasTkAgg( fig , master = root )
        canvas.draw()
        canvas.get_tk_widget().grid( row = 3 , column = 0 , columnspan = 9 , sticky = "nsew" )

def create_layout():
    load_button.grid( row = 0 , column = 0 , padx = 5 , pady = 5 , sticky = "w" )
    sheet_label.grid( row = 0 , column = 1 , padx = 5 , pady = 5 , sticky = "w" )
    sheet_menu.grid( row = 0 , column = 2 , padx = 5 , pady = 5 , sticky = "w" )
    chart_label.grid( row = 0 , column = 3 , padx = 5 , pady = 5 , sticky = "w" )
    chart_menu.grid( row = 0 , column = 4 , padx = 5 , pady = 5 , sticky = "w" )
    x_column_label.grid( row = 0 , column = 5 , padx = 5 , pady = 5 , sticky = "w" )
    x_column_menu.grid( row = 0 , column = 6 , padx = 5 , pady = 5 , sticky = "w" )
    y_column_label.grid( row = 0 , column = 7 , padx = 5 , pady = 5 , sticky = "w" )
    y_column_menu.grid( row = 0 , column = 8 , padx = 5 , pady = 5 , sticky = "w" )
    x_interval_label.grid( row = 1 , column = 0 , padx = 5 , pady = 5 , sticky = "w" )
    x_interval_entry.grid( row = 1 , column = 1 , padx = 5 , pady = 5 , sticky = "w" )
    y_interval_label.grid( row = 1 , column = 2 , padx = 5 , pady = 5 , sticky = "w" )
    y_interval_entry.grid( row = 1 , column = 3 , padx = 5 , pady = 5 , sticky = "w" )
    plot_button.grid( row = 2 , column = 0 , columnspan = 6 , pady = 10 , sticky = "we" )
    database_button.grid( row = 2 , column = 6 , columnspan = 3 , pady = 10 , sticky = "we" )

def save_to_database():
    try:
        # 使用檔名作為資料庫名稱
        db_name = os.path.splitext( file_name )[0] + ".db"
        db_path = os.path.join( os.getcwd() , db_name )
        
        # 如果有同名資料庫就儲存，沒有就建立一個
        conn = sq.connect( db_path )
        for sheet_name , df in data.items():
            df.to_sql( sheet_name , conn , if_exists = "replace" , index = False )
        conn.close()
        messagebox.showinfo( "儲存成功" , f"資料已儲存在資料庫：{db_path}")
    except Exception as e:
        messagebox.showerror( "儲存失敗" )

root = tk.Tk()
root.title( "資料視覺化" )

canvas = None

load_button = tk.Button( root , text = "載入excel檔案" , command = load_excel_file )

sheet_label = tk.Label( root , text = "選擇工作表：" )
selected_sheet = tk.StringVar( root )
sheet_menu = tk.OptionMenu( root , selected_sheet , "" )

chart_label = tk.Label( root , text = "選擇圖表種類：" )
chart_options = ["折線圖" , "點狀圖" , "柱狀圖" , "圓餅圖" , "相關係數"]
selected_chart = tk.StringVar( root )
selected_chart.set( chart_options[0] )
chart_menu = tk.OptionMenu( root , selected_chart , *chart_options )

x_column_label = tk.Label( root , text = "選擇X軸變數：" )
selected_x_column = tk.StringVar(root)
x_column_menu = tk.OptionMenu(root, selected_x_column, "")

y_column_label = tk.Label( root , text = "選擇Y軸變數：" )
selected_y_column = tk.StringVar( root )
y_column_menu = tk.OptionMenu( root , selected_y_column , "" )

x_interval_label = tk.Label( root , text = "X軸單位刻度：" )
x_interval_entry = tk.Entry( root )

y_interval_label = tk.Label( root , text = "Y軸單位刻度：" )
y_interval_entry = tk.Entry( root )

plot_button = tk.Button( root , text = "繪製圖表" , command = plot_selected_chart )

database_button = tk.Button( root , text = "儲存到資料庫" , command = save_to_database )

create_layout()

root.mainloop()
