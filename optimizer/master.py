from optimizer import constant, utils
import datetime

class masterClass:

    def __init__(self) -> None:
        self.route              = { }
        self.customer_route     = { }
        self.routeAssign        = { }
        self.skill              = { }
        self.customer_skill     = { }
        self.overworktime       = { }
        self.overworktimeRaw    = { }
        self.workedtimeRawDaily = { }
        self.holiday            = { }
        self.otherwork          = { }
        self.workedRouteRaw     = { }
        self.worker_code        = { }
        self.worker_name        = { }
        self.otherwork_name     = { }
        self.ignores            = { }
        self.runtime            = { }
        
        pass
    #顧客ルート表とスキル表に基づきルートごとのスキル表を作成
    def build_skill(self, N = 80):

        self.skill = dict()
        
        for worker_code, customer_code_dict in self.customer_skill.items():
            self.skill[worker_code] = dict()
            #print(customer_code_dict)
            for route_code, customer_codes in self.customer_route.items():
                nlength = len(customer_codes)
                counters = [0,0,0,0,0,0,0]
                for customer_code in customer_codes:
                    level = int(customer_code_dict.get(customer_code, 0))
                    counters[level] += 1
                level = 0                
                if(0.01 * N <= sum(counters[2:]) / nlength):
                    level = 10
                    level += counters[3]                    

                elif(0.01 * N <= sum(counters[1:]) / nlength):
                    level = 5
                    if(counters[0]):
                        level = 1
                self.skill[worker_code][route_code] = level
        
        return 
    #マスタの整合性チェック
    def error_check(self, model_params):
                
        nday   : int         = model_params[constant.param_nday]["value"]
        start  : datetime    = model_params[constant.param_targetmonth]["value"]
        
        errMsg = ""        
        # ワーカコードが存在するか
        for worker_name, routes in self.skill.items():
            if not worker_name in self.worker_code:
                errMsg += f"{worker_name} : ワーカコードが存在しませんでした(スキル表と休日の氏名の不一致の可能性あり)\n"            

        # 各ワーカごとに休日の設定があるか？        
        for worker_name, routes in self.skill.items():
            if(not worker_name in self.holiday and not worker_name in self.ignores):
                errMsg += f"{worker_name} : 休日の設定がありませんでした\n"

        # 残業予定にあるか？
        for worker_name, routes in self.skill.items():
            if(not worker_name in self.overworktime):
                errMsg += f"{worker_name} : 残業予定がありませんでした\n"
        
        route_set = set()
        for worker_name, routes in self.skill.items():
            for route_code, level in routes.items():
                route_set.add(route_code)
        
        # ルートの時間設定があるか？        
        for route_code in route_set:
            if(not route_code in self.route):
                errMsg += f"{route_code} : ルート表に存在しませんでした\n"
            else :
                for i in range(-1, nday):
                    date = start + datetime.timedelta(i)
                    if(not date in self.route[route_code]):
                        errMsg += f"{route_code} {date.strftime('%Y/%m/%d')} : ルート表に日付が存在しませんでした\n"
        
        # ルート割り当て表に存在するか？
        for route_code in route_set:
            if(not route_code in self.routeAssign):
                errMsg += f"{route_code} : ルート割当表に存在しませんでした\n"
            else:
                for i in range(nday):
                    date = start + datetime.timedelta(i)
                    if(not date in self.routeAssign[route_code]):
                        errMsg += f"{route_code} {date.strftime('%Y/%m/%d')} : ルート割当表に存在しませんでした\n"
        """
        # 年度開始から残業実績があるか
        fiscal_year_start = utils.get_fiscal_year_start(start)
        for worker_name, routes in self.skill.items():
            if(not worker_name in self.overworktimeRaw):
                errMsg += f"{worker_name} : 残業実績がありませんでした\n"
            elif(not fiscal_year_start in self.overworktimeRaw[worker_name]):
                errMsg += f"{worker_name} {fiscal_year_start} : 残業実績がありませんでした\n"

        # ２週間分の作業実績があるか？
        for worker_name, routes in self.skill.items():
            if(not worker_name in self.worker_code):
                continue

            worker_code = self.worker_code[worker_name]
            if(not worker_code in self.workedtimeRawDaily):
                errMsg += f"{worker_name} {worker_code}: 作業実績がありませんでした\n"
            else:
                for i in range(-13, 0):
                    date = start + datetime.timedelta(i)
                    if(not date in self.workedtimeRawDaily[worker_code]):
                        errMsg += f"{worker_name} {worker_code} {date.strftime('%Y/%m/%d')}: 作業実績がありませんでした\n"
        """
        if(errMsg):
            raise Exception(f"マスタデータに不備がありました : \n{errMsg}")

        return 
