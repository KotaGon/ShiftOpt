#from tkinter import filedialog
from optimizer import optimize
from dateutil.relativedelta import relativedelta
import datetime

"""
#マスタファイルの設定画面を表示
def getFilePathByDialog(*args):

    if(optimize.debug):
        return args;
    
    ret = [ ]
    typ = [('Excel', '*.xlsx')]
    dir = "./"
    for i in range(len(args)):   
        path = ""            
        print(f"ファイル選択してください")
        print(">>", end = "")
        path = filedialog.askopenfilename(filetypes = typ, initialdir = dir);
        print(path)
        if(path == ""):
            raise Exception(FileNotFoundError);
        ret.append(path)    
    return tuple(ret);
"""

def to_month(value : datetime.datetime) -> str:
    return f"{value.year}_{value.month}_1"

def is_numeric(value) -> bool:
    try:
        float(value)
        return True
    except:
        return False

def to_datetime(value : str) -> datetime.datetime:
    for format in ['%Y/%m/%d %H:%M', '%Y-%m-%d %H:%M:%S']:
        try:
            date = datetime.datetime.strptime(value, format)
            return date, True
        except:
            continue
    return None, False

#Fiscal year
def get_fiscal_year_start(date : datetime.datetime) -> datetime.datetime:
    # 日本の年度開始月と日
    fiscal_year_start_month = 4
    fiscal_year_start_day = 1
    
    # 指定された日付の年
    year = date.year
    
    # 4月1日を基準に、指定された日付がその年度の範囲内か判定
    fiscal_year_start = datetime.datetime(year, fiscal_year_start_month, fiscal_year_start_day)
    
    # 指定された日付が4月1日より前の場合、前年度の開始日を返す
    if date < fiscal_year_start:
        fiscal_year_start = datetime.datetime(year-1, fiscal_year_start_month, fiscal_year_start_day)
    
    return fiscal_year_start

#計画期間内の年度内の月を取得
def get_months_in_period(start_date, nday):
    
    #年度はじめの月
    fiscal_year_start = get_fiscal_year_start(start_date)
        
    # 開始日と終了日をdatetimeオブジェクトに変換
    end_date = start_date + datetime.timedelta(nday)
    start_date = start_date.replace(day = 1)
    # 結果を保持するリスト
    months = []
    
    # 開始日から終了日までループ
    current_date = start_date
    #while current_date + relativedelta(months=1) <= end_date:
    while current_date <= end_date:
        # 現在の月をリストに追加                
        if(fiscal_year_start == get_fiscal_year_start(current_date)):
            months.append(current_date)
        # 次の月に移動
        current_date += relativedelta(months=1)
    return months
