import pandas as pd
import matplotlib.pyplot as plt
import os
import sqlite3 as sq
import tkinter as tk
from tkinter import filedialog, messagebox
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def set_chinese_font():
    import matplotlib.font_manager as fm
    font_paths = [
        'C:/Windows/Fonts/msjh.ttc',
        'C:/Windows/Fonts/msjh.ttf',
        'C:/Windows/Fonts/msjhl.ttc',
        'C:/Windows/Fonts/msjhl.ttf',
        'C:/Windows/Fonts/微軟正黑體.ttf'
    ]
    for font_path in font_paths:
        if os.path.exists(font_path):
            msjh_font = fm.FontProperties(fname=font_path)
            plt.rcParams['font.family'] = msjh_font.get_name()
            return
    raise FileNotFoundError("無法找到'微軟正黑體'字體")

def plot_chart(df, chart_type, x_column, y_column, x_interval, y_interval):
    set_chinese_font()
    fig, ax = plt.subplots()

    if x_column == 'Date':
        df[x_column] = pd.to_datetime(df[x_column], format='%m月%d日', errors='coerce')
        df = df.dropna(subset=[x_column])

    if chart_type == "折线图":
        ax.plot(df[x_column], df[y_column], marker='o', label=y_column)
        ax.set_title('Line Chart')
    elif chart_type == "点状图":
        ax.scatter(df[x_column], df[y_column], label=y_column)
        ax.set_title('Scatter Chart')
    elif chart_type == "柱状图":
        ax.bar(df[x_column], df[y_column], label=y_column)
        ax.set_title('Bar Chart')
    elif chart_type == "直方图":
        ax.hist(df[y_column], bins=10, alpha=0.5, label=y_column)
        ax.set_title('Histogram')
    elif chart_type == "圆饼图":
        ax.pie(df[y_column], labels=df[x_column], autopct='%1.1f%%')
        ax.set_title('Pie Chart')

    ax.legend()
    ax.grid(True)

    if x_interval > 0:
        ax.xaxis.set_major_locator(plt.MultipleLocator(x_interval))
    if y_interval > 0:
        ax.yaxis.set_major_locator(plt.MultipleLocator(y_interval))

    return fig

def load_excel_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if not filepath:
        return
    try:
        global data
        global file_name
        global sheet_names

        file_name = os.path.basename(filepath)
        data = pd.read_excel(filepath, sheet_name=None, index_col=0)

        for sheet in data:
            data[sheet] = data[sheet].dropna(how='all')

        sheet_names = list(data.keys())

        sheet_menu['menu'].delete(0, 'end')
        for sheet_name in sheet_names:
            sheet_menu['menu'].add_command(label=sheet_name, command=tk._setit(selected_sheet, sheet_name, update_column_menus))
        selected_sheet.set(sheet_names[0])

        update_column_menus()

        messagebox.showinfo("加载成功", f"成功加载Excel文件：{filepath}")
    except Exception as e:
        messagebox.showerror("加载失败", f"无法加载Excel文件：{str(e)}")

def update_column_menus(*args):
    sheet_name = selected_sheet.get()
    if sheet_name:
        columns = data[sheet_name].columns
        x_column_menu['menu'].delete(0, 'end')
        y_column_menu['menu'].delete(0, 'end')
        for column in columns:
            x_column_menu['menu'].add_command(label=column, command=tk._setit(selected_x_column, column))
            y_column_menu['menu'].add_command(label=column, command=tk._setit(selected_y_column, column))
        selected_x_column.set(columns[0])
        selected_y_column.set(columns[1] if len(columns) > 1 else columns[0])

def plot_selected_chart():
    chart_type = selected_chart.get()
    sheet_name = selected_sheet.get()
    x_column = selected_x_column.get()
    y_column = selected_y_column.get()
    x_interval = float(x_interval_entry.get()) if x_interval_entry.get() else 1
    y_interval = float(y_interval_entry.get()) if y_interval_entry.get() else 1

    df = data[sheet_name]

    # print(f"选中的工作表: {sheet_name}")
    # print(f"选中的X轴数据列: {x_column}")
    # print(f"选中的Y轴数据列: {y_column}")
    # print(f"数据预览:\n{df[[x_column, y_column]].head()}")
    # print(f"数据类型:\n{df.dtypes}")

    fig = plot_chart(df, chart_type, x_column, y_column, x_interval, y_interval)
    if fig:
        global canvas
        if canvas:
            canvas.get_tk_widget().grid_forget()
        canvas = FigureCanvasTkAgg(fig, master=root)
        canvas.draw()
        canvas.get_tk_widget().grid(row=3, column=0, columnspan=9, sticky='nsew')

def create_layout():
    load_button.grid(row=0, column=0, padx=5, pady=5, sticky='w')
    sheet_label.grid(row=0, column=1, padx=5, pady=5, sticky='w')
    sheet_menu.grid(row=0, column=2, padx=5, pady=5, sticky='w')
    chart_label.grid(row=0, column=3, padx=5, pady=5, sticky='w')
    chart_menu.grid(row=0, column=4, padx=5, pady=5, sticky='w')
    x_column_label.grid(row=0, column=5, padx=5, pady=5, sticky='w')
    x_column_menu.grid(row=0, column=6, padx=5, pady=5, sticky='w')
    y_column_label.grid(row=0, column=7, padx=5, pady=5, sticky='w')
    y_column_menu.grid(row=0, column=8, padx=5, pady=5, sticky='w')
    x_interval_label.grid(row=1, column=0, padx=5, pady=5, sticky='w')
    x_interval_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
    y_interval_label.grid(row=1, column=2, padx=5, pady=5, sticky='w')
    y_interval_entry.grid(row=1, column=3, padx=5, pady=5, sticky='w')
    plot_button.grid(row=2, column=0, columnspan=6, pady=10, sticky='we')

root = tk.Tk()
root.title("Excel 数据可视化")

canvas = None

load_button = tk.Button(root, text="加载Excel文件", command=load_excel_file)

sheet_label = tk.Label(root, text="选择工作表：")
selected_sheet = tk.StringVar(root)
sheet_menu = tk.OptionMenu(root, selected_sheet, "")

chart_label = tk.Label(root, text="选择图表类型：")
chart_options = ["折线图", "点状图", "柱状图", "直方图", "圆饼图"]
selected_chart = tk.StringVar(root)
selected_chart.set(chart_options[0])
chart_menu = tk.OptionMenu(root, selected_chart, *chart_options)

x_column_label = tk.Label(root, text="选择X轴数据列：")
selected_x_column = tk.StringVar(root)
x_column_menu = tk.OptionMenu(root, selected_x_column, "")

y_column_label = tk.Label(root, text="选择Y轴数据列：")
selected_y_column = tk.StringVar(root)
y_column_menu = tk.OptionMenu(root, selected_y_column, "")

x_interval_label = tk.Label(root, text="设置X轴刻度间隔：")
x_interval_entry = tk.Entry(root)

y_interval_label = tk.Label(root, text="设置Y轴刻度间隔：")
y_interval_entry = tk.Entry(root)

plot_button = tk.Button(root, text="绘制图表", command=plot_selected_chart)

create_layout()

root.mainloop()
