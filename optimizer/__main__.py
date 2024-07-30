from optimizer import optimize
from optimizer import constant
from optimizer import utils
from optimizer.master import masterClass
from ortools.sat.python import cp_model
from collections import defaultdict
import warnings
import pandas as pd
import argparse
import datetime
import sys
import logging

#設定ファイル読み込み
def read_params(sht):
    try:
        setting_params = {
            constant.param_outputfile                  : { "type" : "string",   "value" : "./data/result.xlsx"},
            constant.param_targetmonth                 : { "type" : "datetime", "value" : None},
            constant.param_nday                        : { "type" : "int",      "value" : 0},
            constant.param_timelimit                   : { "type" : "float",    "value" : 120}, 
            constant.param_publicholiday               : { "type" : "int",      "value" : 9},
            constant.param_skill_standard_val          : { "type" : "int",      "value" : 80},
            constant.param_holiday_const_name          : { "type" : "int",   "value" : 0},
            constant.param_pulic_holiday_const_name    : { "type" : "int",   "value" : 0},
            constant.param_worktime_year_const_name    : { "type" : "int",   "value" : 0},
            constant.param_worktime_month_const_name   : { "type" : "int",   "value" : 0},
            constant.param_resttime_const_name         : { "type" : "int",   "value" : 0},
            constant.param_dayave_const_name           : { "type" : "int",   "value" : 0},
            constant.param_weekave_const_name          : { "type" : "int",   "value" : 0},         
            constant.param_department                  : { "type" : "string", "value" : ""},   
            constant.param_diffstarttime               : { "type" : "int",   "value" : 120}   
            
        }
         
        nrow = len(sht)
        for i in range(nrow):
            key, val = sht.iloc[i,0], sht.iloc[i,1]
            if(pd.isnull(val) or pd.isna(val)):
                continue
            elif(not key in setting_params):
                continue
            type = setting_params[key]["type"]
            if(type == "string"):
                setting_params[key]["value"] = val
            elif(type == "datetime"):
                setting_params[key]["value"], good = utils.to_datetime(val)    
                if(not good):
                    raise(Exception("datetimeの変換失敗"))
            elif(type == "int"):  
                if(utils.is_numeric(val)):
                    setting_params[key]["value"] = int(val)
                else:
                    raise(Exception("数値の変換失敗"))
            elif(type == "float"):
                if(utils.is_numeric(val)):
                    setting_params[key]["value"] = float(val)
                else:
                    raise(Exception("数値の変換失敗"))
        
    except Exception as e:
        print(" × 設定読み込み失敗")
        exit()

    print(" ◯ 設定読み込み完了")

    return setting_params

#ルート表読み込み
def read_route(sht):
    route, run_time = defaultdict(dict), defaultdict(dict)
    try:
        nrow = len(sht)
        for i in range(2, nrow):
            date, good = utils.to_datetime(sht.iloc[i, 0])
            route_code = sht.iloc[i, 1]
            startime, endtime = str(sht.iloc[i, 32]), str(sht.iloc[i, 35])
            startime, endtime = startime.split(":"), endtime.split(":")
            route[route_code][date] = (startime, endtime)

            rtime = int(sht.iloc[i, 36])
            run_time[route_code][date] = rtime 

    except Exception as e:
        print(" × ルート表読み込み失敗")
        exit()
    print(" ◯ ルート表読み込み完了")
    return route, run_time

#ルート割当表読み込み        
def read_route_assingment(sht):
    ret = dict()
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)

        for i in range(nrow):
            route_code = str(sht.iloc[i, 0])
            ret[route_code] = dict()

            for j in range(1, ncol):
                date, flag = sht.columns[j], sht.iloc[i, j]
                if(pd.isna(flag) or pd.isnull(flag)):
                    flag = 0
                elif(flag == "0" or flag == 0):
                    flag = 0
                elif(flag == "1" or flag == 1):
                    flag = 1
                else:
                    errMsg += f"0,1,空白以外の設定がありました:{i+1}行{j+1}列目 値 {flag}\n"
                ret[route_code][date] = flag    
        
        if(errMsg != ""):
            raise(Exception(errMsg))

    except Exception as e:
        print(" × ルート割当表読み込み失敗")
        print(e)
        exit()
    print(" ◯ ルート割当表読み込み完了")
    return ret

#顧客ルート表読み込み
def read_customer_route(sht):

    ret = dict()
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        
        for i in range(nrow):
            route_code, customer_code = str(sht.iloc[i, 0]), str(sht.iloc[i, 1])
            if(not route_code in ret):
                ret[route_code] = list()
            ret[route_code].append(customer_code)
        if(errMsg != ""):
            raise(Exception(errMsg))
        
    except Exception as e:
        print(" × 顧客ルート表読み込み失敗")
        print(e)
        exit()
    print(" ◯ 顧客ルート表読み込み完了")
    return ret

#スキル表読み込み
def read_skill(sht):
    ret = dict()
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        
        for i in range(nrow):
            name = str(sht.iloc[i, 0])
            ret[name] = dict()
            
            for j in range(1, ncol):
                #route, skill = sht.columns[j], str(sht.iloc[i, j])
                route, skill = str(sht.columns[j]), sht.iloc[i, j]
                
                if(pd.isna(skill) or pd.isnull(skill)):
                    skill = 0
                elif(utils.is_numeric(skill)):
                #elif(pd.to_numeric(skill).notna()):
                    skill = int(sht.iloc[i, j])
                else:
                    errMsg += f"数字以外の設定がありました:{i+1}行{j+1}列目 値 {skill}\n"
                    #raise Exception(f"数字以外の設定がありました:{i+1}行{j+1}列目")
                ret[name][route] = skill
        if(errMsg != ""):
            raise(Exception(errMsg))
        
    except Exception as e:
        print(" × スキル表読み込み失敗")
        print(e)
        exit()
    print(" ◯ スキル表読み込み完了")
    return ret

#勤務時間
def read_overworktime(sht):
    ret = dict()
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        
        for i in range(nrow):
            name = str(sht.iloc[i, 0]).replace(" ", "")
            ret[name] = dict()

            for j in range(1, ncol-2):
                #month, wtime = sht.columns[j], str(sht.iloc[i, j])
                month, good = utils.to_datetime(str(sht.columns[j]))
                wtime = sht.iloc[i, j]

                if(not good):
                    continue

                if(pd.isna(wtime) or pd.isnull(wtime)):
                    wtime = 0
                elif(utils.is_numeric(wtime)):
                    wtime = int(sht.iloc[i, j])
                else:
                    errMsg += f"数字以外の設定がありました:{i+1}行{j+1}列目 値 {wtime}\n"
                    #raise Exception(f"数字以外の設定がありました:{i+1}行{j+1}列目")
                ret[name][month] = wtime
        if(errMsg != ""):
            raise(Exception(errMsg))
    except Exception as e:
        print(" × 残業予定読み込み失敗")
        print(e)
        exit()
    print(" ◯ 残業予定読み込み完了")
    return ret

#休日予定読み込み
def read_holiday(sht):
    worker_names, worker_codes, holidays, otherworks, otherworks_names = dict(), dict(), dict(), dict(), set()

    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        for i in range(7, len(sht.columns)):
            value = sht.iloc[2, i]
            if(pd.isna(value) or pd.isnull(value) or not utils.is_numeric(value)):
                break
            ncol = i + 1
        for i in range(4, nrow):
            #code, name = sht.iloc[i, 4], sht.iloc[i, 5]
            code, name = sht.iloc[i, 4], sht.iloc[i, 5]

            if(not (pd.isna(code) or pd.isnull(code))):
                code = code.replace(" ", "")
                if(not (pd.isna(code) or pd.isnull(code))):
                    worker_codes[name] = code
                    worker_names[code] = name
                holidays[code], otherworks[code] = dict(), dict()

                for j in range(7, ncol):

                    date, good = utils.to_datetime(str(sht.iloc[2, j]))
                    cell = sht.iloc[i, j]
                    if(not good):
                        continue

                    # 作業日
                    if(pd.isna(cell) or pd.isnull(cell) or utils.is_numeric(cell)):
                        #ret[code][date] = 0
                        holidays[code][date] = "" 
                        otherworks[code][date] = ""
                    else:
                        if(str(cell) == "／" or str(cell) == "有休"):
                            #ret[code][date] = 1
                            holidays[code][date] = str(cell)
                            otherworks[code][date] = ""
                        else:
                            #ret[code][date] = 2
                            holidays[code][date] = ""
                            otherworks[code][date] = str(cell)
                            otherworks_names.add(str(cell))
    except Exception as e:
        print(" × 休日予定読み込み失敗")
        print(e)
        exit()
    print(" ◯ 休日予定読み込み完了")
    return worker_names, worker_codes, holidays, otherworks, otherworks_names

#作業実績読み込み
def read_workedroute_raw(sht):
    ret = dict()

    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        
        for i in range(nrow):
            name = str(sht.iloc[i, 0]).replace(" ", "")
            ret[name] = dict()
            
            for j in range(1, ncol):                
                date, raw = sht.columns[j], sht.iloc[i, j]
                date, good = utils.to_datetime(str(date))
                if(not good):
                    continue
                if(pd.isna(raw) or pd.isnull(raw)):
                    raw = ""
                ret[name][date] = raw    
    except Exception as e:
        print(" × 作業ルート実績読み込み失敗")
        print(e)
        exit()
    print(" ◯ 作業ルート実績読み込み完了")
    return ret

#残業実績読み込み
def read_overworktime_raw(sht):
    ret = dict()
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)
        for i in range(nrow):
            name = str(sht.iloc[i, 0]).replace(" ", "")
            ret[name] = dict()
            for j in range(1, ncol):
                month, wtime = sht.columns[j], sht.iloc[i, j]
                month, good = utils.to_datetime(str(month))
                if(not good):
                    continue
                if(pd.isna(wtime) or pd.isnull(wtime)):
                    wtime = 0
                #elif(wtime.isdecimal()):
                elif(utils.is_numeric(wtime)):
                    wtime = int(sht.iloc[i, j])
                else:
                    errMsg += f"数字以外の設定がありました:{i+1}行{j+1}列目 値{wtime}\n"
                    #raise Exception(f"数字以外の設定がありました:{i+1}行{j+1}列目")
                ret[name][month] = wtime
        if(errMsg != ""):
            raise(Exception(e))
    except Exception as e:
        print(" × 残業実績読み込み失敗")
        print(e)
        exit()
    print(" ◯ 残業実績読み込み完了")
    
    return ret

#作業時間実績読み込み
def read_workedtime_dailyraw(sht):
    workedtimeRawDaily, workedRouteRaw,overworktimeRaw = defaultdict(dict), defaultdict(dict), defaultdict(dict)
    
    errMsg = ""
    try:
        nrow = len(sht)
        ncol = len(sht.columns)

        def extract_date(date):
            date = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
            return date.strftime("%Y-%m-%d")
        def extract_time(date):
            try:
                date = datetime.datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
                return date.strftime("%H:%M:%S")
            except:
                return date

        for i in range(1, nrow):
            date, route_code, code, stime, etime = sht.iloc[i, 0], sht.iloc[i, 3], sht.iloc[i, 16], sht.iloc[i, 14], sht.iloc[i, 15]
            
            date, stime, etime = extract_date(date), extract_time(stime), extract_time(etime)
            sdate, edate = date + " " + stime, date + " " + etime

            date = datetime.datetime.strptime(date, "%Y-%m-%d")
            sdate = datetime.datetime.strptime(sdate, "%Y-%m-%d %H:%M:%S")
            edate = datetime.datetime.strptime(edate, "%Y-%m-%d %H:%M:%S")
            diff = edate - sdate

            first_day_of_month = sdate.replace(day=1, hour=0, minute=0, second=0)
            
            # 分単位で取得
            diff_minutes = diff.total_seconds() / 60            
            if(not first_day_of_month in overworktimeRaw[code]):
                overworktimeRaw[code][first_day_of_month] = 0                        
            overworktimeRaw[code][first_day_of_month] += max(0, diff_minutes - 9 * 60)
            
            
            if(not code in workedtimeRawDaily):
                workedtimeRawDaily[code] = dict()
            if(not date in workedtimeRawDaily[code]):
                workedtimeRawDaily[code][date] = [datetime.datetime.max, datetime.datetime.min]
 
            workedtimeRawDaily[code][date][0], workedtimeRawDaily[code][date][1] = sdate, edate
            workedRouteRaw[code][date] = route_code
            #workedtimeRawDaily[code][date][0], workedtimeRawDaily[code][date][1] = min(workedtimeRawDaily[code][date][0], sdate), max(workedtimeRawDaily[code][date][1], edate)

        """
        for i in range(nrow):
            date, code, updatetime = sht.iloc[i, 0], sht.iloc[i, 1], sht.iloc[i, 9]

            if(pd.isna(date) or pd.isnull(date)):
                continue
            if(pd.isna(updatetime) or pd.isnull(updatetime)):
                continue
            date, code, updatetime = str(date), str(code), str(updatetime)

            current_year, current_month, current_day = int(date[0:4]), int(date[4:6]), int(date[6:8])
            date = datetime.datetime(year=current_year, month=current_month, day=current_day)
            updatetime, good = utils.to_datetime(updatetime)
            if(not good):
                continue

            if(not code in ret):
                ret[code] = dict()
            if(not date in ret[code]):
                ret[code][date] = [datetime.datetime.max, datetime.datetime.min]
            ret[code][date][0], ret[code][date][1] = min(ret[code][date][0], updatetime), max(ret[code][date][1], updatetime)

        """

        if(errMsg != ""):
            raise(Exception(errMsg))
    except Exception as e:
        print(" × 作業時間実績読み込み失敗")
        print(e)
        exit()
    print(" ◯ 作業時間実績読み込み完了")
    
    return workedtimeRawDaily, workedRouteRaw, overworktimeRaw

#シフト割り振り無効対象読み込み
def read_ingore_workers_raw(sht):
    ret = set()
    try:
        nrow = len(sht)        
        for i in range(nrow):            
            name = str(sht.iloc[i, 1])
            ret.add(name)
    except Exception as e:
        print(" × シフト割り振り無効対象読み込み失敗")
        print(e)
        exit()
    print(" ◯ シフト割り振り無効対象読み込み完了")
    return ret

#マスタファイル読み込み
def import_data():
    warnings.simplefilter('ignore')
    print("マスタファイル？", end = "")
    #filepath, = utils.getFilePathByDialog("./config/set.xlsx")
    filepath = constant.attr_config + "/" + constant.attr_masterfile
    print(f" → {filepath}")
    excel = pd.read_excel(filepath, sheet_name = None, dtype=str)

    params, master = None, masterClass()

    print("マスタファイル読み込み開始 ...")

    for sht_name in excel:
        #excel[sht_name] = excel[sht_name].replace(" ", "", regex = True )
        #excel[sht_name] = excel[sht_name]

        if(sht_name == constant.attr_setting_shtname):
            params = read_params(excel[sht_name])
        elif(sht_name == constant.attr_route_shtname):
            master.route, master.runtime = read_route(excel[sht_name])
        elif(sht_name == constant.attr_skill_shtname): #スキル表
            master.customer_skill = read_skill(excel[sht_name])
        elif(sht_name == constant.attr_customer_route_shtname): #顧客ルート表
            master.customer_route = read_customer_route(excel[sht_name])
        #    master.skill = read_skill(excel[sht_name])
        elif(sht_name == constant.attr_overworktime_shtname):
            master.overworktime = read_overworktime(excel[sht_name])
        elif(sht_name == constant.attr_holiday_shtname):
            master.worker_name, master.worker_code, master.holiday, master.otherwork, master.otherwork_name = read_holiday(excel[sht_name])
        elif(sht_name == constant.attr_worked_time_raw_shtname):
            master.workedtimeRawDaily, master.workedRouteRaw, master.overworktimeRaw = read_workedtime_dailyraw(excel[sht_name])        
        elif(sht_name == constant.attr_route_assignment_shtname):
            master.routeAssign = read_route_assingment(excel[sht_name])
        elif(sht_name == constant.attr_ingore_workers_shtname):
            master.ignores = read_ingore_workers_raw(excel[sht_name])
        #elif(sht_name == constant.attr_overworktime_raw_shtname):
        #    master.overworktimeRaw = read_overworktime_raw(excel[sht_name])
        #elif(sht_name == constant.attr_worked_route_raw_shtname):
        #    master.workedRouteRaw = read_workedroute_raw(excel[sht_name])
        
    print("")
    print("マスタファイル読み込み完了")    

    master.error_check(params)    
    master.build_skill(params[constant.param_skill_standard_val]["value"])

    #print(master.skill)
    #input()

    return params, master

#ロガー
def getLogger() -> logging.Logger:
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)

    # ファイルハンドラ作成
    #formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = logging.FileHandler(constant.attr_config + "/" + constant.attr_log_file, mode='a', encoding='utf-8')
    #file_handler = logging.FileHandler(constant.attr_config + "/" + constant.attr_log_file)
    file_handler.setLevel(logging.DEBUG)
    #file_handler.setFormatter(formatter)

	# ロガーに追加
    logger.addHandler(file_handler)
    return logger

#実行
def run(params : dict, master : dict, logger : logging.Logger) -> None:
    solver = optimize.solverClass(params, master, logger)
    status = solver.optimize()

    outputfile = params[constant.param_outputfile]["value"]    
    if(status == cp_model.FEASIBLE or status == cp_model.OPTIMAL) :
        solver.logger.info("最適化完了")
        solver.output(outputfile)
    else:
        solver.logger.info("最適化失敗")

    return None

if __name__ == "__main__":
    logger = getLogger()
    parser = argparse.ArgumentParser()

    parser.add_argument("-d", "--debug", action="store_true", help="debug flag")
    args = parser.parse_args()
    params, master = import_data()
    run(params, master, logger)
    """
    try:
        parser.add_argument("-d", "--debug", action="store_true", help="debug flag")
        args = parser.parse_args()
        params, master = import_data()
        run(params, master, logger)

    except Exception as e:
        print("エラー発生しました。")
        print(e)
        logger.error(f"エラー発生:{datetime.datetime.now().strftime('%m/%d %H:%M:%S')}")
        logger.exception(sys.exc_info())
        print(sys.exc_info())
        pass
    """
