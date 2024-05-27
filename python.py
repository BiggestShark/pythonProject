import pandas as pd ## 用於讀取檔案
import matplotlib.pyplot as plt ## 用於製作圖表
import os ## 用於獲取檔案名稱
import sqlite3 as sq ## 資料庫

def line_chart( data ):
    for column in data.columns[1:]:
        plt.plot( data[data.columns[0]] , data[column] , marker = "o" , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( "折線圖" )
    plt.legend()
    plt.grid( True )
    plt.show()

def scatter_chart( data ):
    for column in data.columns[1:]:
        plt.scatter( data[data.columns[0]] , data[column] , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( '點狀圖' )
    plt.legend()
    plt.grid(True)
    plt.show()

def bar_chart( data ):
    for column in data.columns[1:]:
        plt.bar( data[data.columns[0]] , data[column] , label = column )
    plt.xlabel( data.columns[0] )
    plt.ylabel( data.columns[1] )
    plt.title( "柱狀圖" )
    plt.legend()
    plt.grid(True)
    plt.show()

def histogram( data ):
    for column in data.columns[1:]:
        plt.hist( data[column] , bins=10 , alpha = 0.5 , label = column )
    plt.xlabel('Value')
    plt.ylabel('Frequency')
    plt.title('Histogram')
    plt.legend()
    plt.grid(True)
    plt.show()

def pie_chart( data ):
    # 使用第一列數據
    if len( data.columns ) > 1:
        plt.pie( data[data.columns[1]] , labels = data[data.columns[0]] , autopct = "%1.1f%%" )
        plt.title( "圓餅圖" )
        plt.show()
    else:
        print( "Data not sufficient for pie chart" )

db_path = "C:/Users/User/Documents/Python Scripts/sqlite"

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

db_name = os.path.splitext( file_name )[0] + '.db' ## 設定資料庫名稱
conn = sq.connect( db_name ) ## 連接資料庫
for sheet_name , df in data.items():
    df.to_sql( sheet_name , conn , if_exists = 'replace' , index = False )

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

conn.close()