# coding = utf-8
import win32api
import win32gui
import win32con
import win32clipboard as clipboard
import time
# from PIL import Image
# from io import BytesIO
import requests
from lxml import etree
import pymysql
import argparse
import sys
import datetime
import random

# python C:\Users\Siqi\Desktop\OptionFlow\broadcast_local_one_session_prod.py -url https://app.flowalgo.com/ -login_url https://app.flowalgo.com/users/login -proxy http:5.79.66.2:13010 -username option20201216 -password option123 -free_target "" -vip_target "Option Flow VIP" -million_target "百万大单"
# python C:\Users\Siqi\Desktop\OptionFlow\broadcast_local_one_session_prod.py -url https://app.flowalgo.com/ -login_url https://app.flowalgo.com/users/login -proxy http:5.79.66.2:13010 -username option20201216 -password option123 -free_target "" -vip_target "Option Flow VIP,Option Flow 2 群" -million_target "百万大单"


parser = argparse.ArgumentParser()
parser.add_argument('-url', type=str, default=None)
parser.add_argument('-login_url', type=str, default=None)
parser.add_argument('-proxy', type=str, default=None)
parser.add_argument('-username', type=str, default=None)
parser.add_argument('-password', type=str, default=None)
parser.add_argument('-free_target', type=str, default=None)
parser.add_argument('-vip_target', type=str, default=None)
parser.add_argument('-million_target', type=str, default=None)
args = parser.parse_args()

# headers = {
# # firefox
# 'user-agent':''
# ,'cookie':''.encode('utf-8')
# ,'Connection':'keep-alive'
#
# # chrome
# # 'user-agent':''
# # ,'cookie':''
# # ,'Connection':'keep-alive'
# }

# headers = {
# 'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
# ,'cookie':'__cfduid=df32cb19205bc7e4b8ba14a3ccd5d96cf1601614085; _ga=GA1.2.1122056143.1601614090; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; intercom-id-dtoll8e6=5f313d1c-1f95-4882-8841-ca87260d4020; _fbp=fb.1.1601614094994.719174546; _gid=GA1.2.213395945.1601871573; _gat_gtag_UA_105239038_2=1; PHPSESSID=02e58170f78d7c4cabc492703794ad83; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2015595%7D; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201001%3A4%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201001%3A4%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201001%3A4; intercom-session-dtoll8e6=VmZkSmVjNHlxVHdLREllNnNMWms1WkszekFQYmhkK2VxUmdkd3ZxaGRZMlZsRTVoSjZnTCt3NGZlTmMyWUZtSy0tbUJZMmlHWlM0UjVRNGFLaUpKMjlzdz09--e1239d20d39fe61462ccd4ecc1e48347a35076a4; amember_nr=3af6026189558ad8013941885584265a; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=audreyb123%7C1602044398%7CGT7BSbkNM73utbELpagSfz7KJ5DHqPkUezQe0XArL59%7C7a8d4e15aa22babe048f2b00d7056a50ff45d6a0c5cf5038f076f8909fc34bbf'
# }

headers = {
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
,'cookie':'_ga=GA1.2.1122056143.1601614090; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; intercom-id-dtoll8e6=5f313d1c-1f95-4882-8841-ca87260d4020; _fbp=fb.1.1601614094994.719174546; __adroll_fpc=ce481f124e8b1fd7a3839ad834766b81-1601960955579; __cfduid=db83a7f30ccada41ea95ef04c955224281606975395; _gid=GA1.2.2033686401.1608101444; amember_nr=57634ba67e67e0ce917eeff185137924; _gat_gtag_UA_105239038_2=1; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20210015%3A2%7CJURQBX5ZWNGNLAFB4ISSVP%3A20210015%3A2%7CEWEJP57Y6NEUVKJVXOCEJP%3A20210015%3A2; PHPSESSID=bb7275d8646b9fe506a032961f134149; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201216%7C1608441079%7CjOBtG6ahyXYt2s5gHdYTeG5NX3spf91i7bMUQCyZshd%7C3f64af3a72d2d734dc607f311bb9fca139978c39b5211cf0d569ac87fcac2bce; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2018571%2C%22%24device_id%22%3A%20%22174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2018571%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%7D; intercom-session-dtoll8e6=ekpPa0VSMDNyN0FoTngwenFCRlk5OTVka0ZHMVlURFFkZjZMRVAwTDQyOFNFYmtNUlhSYkRDd0pBUnpZdHBvay0temhvL1NmY3BYd1dFQk0rKzNSaUpmdz09--c7529533a27325045449d7f97a62a63299fe9fa7; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20210015%3A3%7CJURQBX5ZWNGNLAFB4ISSVP%3A20210015%3A2%7CEWEJP57Y6NEUVKJVXOCEJP%3A20210015%3A2'
}


running_count = 1

time_point = max(9, datetime.datetime.now().hour+1 if datetime.datetime.now().hour+1<=16 else 10)
# time_point = 23


def build_database_connection():
    connection = pymysql.connect(
        host="127.0.0.1",
        user='lfang',
        password='123456',
        database='FlowAlgo_dev',
        port=3306
    )
    return connection


def get_data(sql):
    conn = build_database_connection()

    # print('Database connection built!')

    cursor = conn.cursor()

    cursor.execute(sql)

    results = cursor.fetchall()

    print('Data returned!')

    conn.close()

    return results


def telegram_bot_sendtext(chatID, bot_message):
    print('execute telegram_bot_sendtext')
    bot_token = '1323919359:AAEtt77oSWn4rHxExKvSB3QDmyZu9hnGblM'
    bot_chatID = chatID
    send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=Markdown&text=' + bot_message

    response = requests.get(send_text)

    return response.json()


class Option:
    def __init__(self):
        self.url = args.url
        self.login_url = args.login_url
        self.proxy = {args.proxy.split(':')[0]: args.proxy.split(':')[1] + ':' + args.proxy.split(':')[2]}
        self.username = args.username
        self.password = args.password
        self.free_target = args.free_target.split(',')
        self.vip_target = args.vip_target.split(',')
        self.million_target = args.million_target.split(',')

    def get_session(self):
        data = {
            'amember_login': self.username
            , 'amember_pass': self.password
        }
        try:
            session = requests.session()
            session.post(self.login_url, data)
            print('Session created!')
            return session
        except:
            telegram_bot_sendtext('-408542611', "Option flow job failed at {}".format(str(datetime.datetime.now())))
            return None

    def extract(self, session):
        try:
            response = session.get(self.url, headers=headers)
        except:
            telegram_bot_sendtext('-408542611', "Option flow job failed at {}".format(str(datetime.datetime.now())))
            return

        parser = etree.HTMLParser(encoding='utf-8')

        htmlElement = etree.fromstring(response.content.decode('utf-8'), parser=parser)

        data = htmlElement.xpath('//*[@id="optionflow"]/div[2]//div[@class and @data-ticker and @data-sentiment and @data-flowid and @data-premiumpaid and @data-ordertype]')
        print('Number of records', len(data))
        if len(data) == 0:
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            print('='*100)
            # print(response.content.decode('utf-8'))
            telegram_bot_sendtext('-408542611', 'No records return from website!!!')

        conn = build_database_connection()

        cursor = conn.cursor()

        sql_truncate_realtime = """truncate table realtime_option_flow;"""

        cursor.execute(sql_truncate_realtime)

        for d in data:
            # time
            # time = d.xpath('./div[1]//text()')[0]
            time = d.xpath("./div[@class='time']/span/text()")[0]
            timestamp = d.xpath("./div[@class='time']/@data-time")[0]

            # ticker
            # ticker = d.xpath('./div[2]//text()')[0]
            ticker = d.xpath("./div[@class='ticker']/span/text()")[0]

            # expiry
            # expiry = d.xpath('./div[3]/span//text()')[0]
            expiry = d.xpath("./div[@class='expiry']/span/text()")[0]

            # strike
            # strike = d.xpath('./div[4]/span//text()')[0]
            strike = d.xpath("./div[@class='strike']/span/text()")[0]

            # contract - type
            # ContractType = d.xpath('./div[5]/span//text()')[0] +d.xpath('./div[5]/span//text()')[1]
            ContractType = d.xpath("./div[@class='contract-type']/span//text()")[0] \
                           + d.xpath("./div[@class='contract-type']/span//text()")[1]

            # ref - Spot Price
            # SpotPrice = d.xpath('./div[6]/span//text()')[0]
            SpotPrice = d.xpath("./div[@class='ref']/span/text()")[0]

            # Contract Size | Price
            # ContractSizePrice = d.xpath('./div[7]/span//text()')[0]
            ContractSizePrice = d.xpath("./div[@class='details']/span/text()")[0]

            # Multi Exchange Sweep
            # ContractSizePrice = d.xpath('./div[8]/span/span//text()')[0] + d.xpath('./div[8]/span/span//text()')[1]
            MultiExchangeSweep = d.xpath("./div[@class='type']/span/span//text()")[0] \
                                 + d.xpath("./div[@class='type']/span/span//text()")[1]

            # premium
            # premium = d.xpath('./div[9]/span//text()')[0]
            premium = d.xpath("./div[@class='premium']/span/text()")[0]

            # sector
            # sector = d.xpath('./div[12]/span//text()')[0]
            sector = d.xpath("./div[@class='sector']/span/text()")[0]

            # magnitude
            # //div[@class='magnitude']/span/i/@title/
            magnitude = d.xpath("./div[@class='magnitude']/span/i/@title")[0]

            sql = """
                insert into realtime_option_flow(transcation_time,transcation_timestamp,ticker,expiry,ContractType,strike,SpotPrice,ContractSizePrice,MultiExchangeSweep,premium,sector,magnitude,row_hash)
                values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,md5(%s));
                """
            cursor.execute(sql, (
            time, timestamp, ticker, expiry, ContractType, strike, SpotPrice, ContractSizePrice, MultiExchangeSweep,
            premium, sector, magnitude,
            time + timestamp + ticker + expiry + ContractType + strike + SpotPrice + ContractSizePrice + MultiExchangeSweep + premium + sector + magnitude))

        conn.commit()

        print('Data extracted!')

        conn.close()

    @staticmethod
    def remove_duplicates():
        conn = build_database_connection()

        cursor = conn.cursor()

        sql_truncate_new = """truncate table new_option_flow;"""

        cursor.execute(sql_truncate_new)

        sql_insert_new = """
        INSERT INTO new_option_flow (transcation_time,ticker,expiry,ContractType,strike,SpotPrice,ContractSizePrice,MultiExchangeSweep,premium,premium_quantity,sector,magnitude,row_hash,transcation_timestamp)
        SELECT   
        rt.transcation_time,
        rt.ticker,
        STR_TO_DATE(rt.expiry, '%m/%d/%Y'),
        rt.ContractType,
        rt.strike*1.0,
        rt.SpotPrice*1.0,
        rt.ContractSizePrice,
        rt.MultiExchangeSweep,
        rt.premium,
        case 
        when rt.premium like '%K%' then TRIM(TRAILING 'K' FROM TRIM(LEADING '$' from rt.premium))*1000 
        when rt.premium like '%M%' then TRIM(TRAILING 'M' FROM TRIM(LEADING '$' from rt.premium))*1000000
        end,
        rt.sector,
        rt.magnitude*1,
        rt.row_hash,
        case 
        	when rt.transcation_timestamp = '' then STR_TO_DATE(concat(DATE_FORMAT(NOW(), '%Y-%m-%d'),' ', rt.transcation_time), '%Y-%m-%d %h:%i %p')
        	else FROM_UNIXTIME(rt.transcation_timestamp * 1)- INTERVAL 3 hour
        end
        FROM realtime_option_flow rt
        left join FlowAlgo_dev.historical_option_flow his on rt.row_hash = his.row_hash
        WHERE his.row_hash is NULL; 
        """

        cursor.execute(sql_insert_new)
        print('Insert into new finish')

        sql_truncate_realtime = """
        truncate table realtime_option_flow;
        """

        cursor.execute(sql_truncate_realtime)
        print('truncate realtime finish')

        sql_insert_historical = """
        INSERT INTO historical_option_flow (transcation_time,ticker,expiry,ContractType,strike,SpotPrice,ContractSizePrice,MultiExchangeSweep,premium,premium_quantity,sector,magnitude,row_hash,transcation_timestamp,original_etl_insert_time)
        SELECT transcation_time,ticker,expiry,ContractType,strike,SpotPrice,ContractSizePrice,MultiExchangeSweep,premium,premium_quantity,sector,magnitude,row_hash,transcation_timestamp,original_etl_insert_time  
        FROM new_option_flow;
        """

        cursor.execute(sql_insert_historical)
        print('Insert into historical finish')

        conn.commit()

        conn.close()

    @staticmethod
    def send_msg(win):
        # 以下为“CTRL+V”组合键,回车发送，（方法一）
        win32api.keybd_event(17, 0, 0, 0)  # 有效，按下CTRL
        time.sleep(1)  # 需要延时
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, 86, 0)  # V
        win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # 放开CTRL
        time.sleep(1)  # 缓冲时间
        win32gui.SendMessage(win, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)  # 回车发送
        return

    @staticmethod
    def txt_ctrl_v(txt_str):
        # 定义文本信息,将信息缓存入剪贴板
        clipboard.OpenClipboard()
        clipboard.EmptyClipboard()
        clipboard.SetClipboardData(win32con.CF_UNICODETEXT, txt_str)
        clipboard.CloseClipboard()
        return

    @staticmethod
    def local_win(title_name):
        win = win32gui.FindWindow('ChatWnd', title_name)
        print("找到句柄：%x" % win)
        if win != 0:
            left, top, right, bottom = win32gui.GetWindowRect(win)
            print(left, top, right, bottom)  # 最小化为负数
            # 最小化时点击还原，下面为单个窗口
            if top < 0:
                # 鼠标点击，还原窗口
                win32api.SetCursorPos([190, 1040])  # 鼠标定位到(190,1040)
                # 执行左单键击，若需要双击则延时几毫秒再点击一次即可
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                ######点击完成一次
            time.sleep(1)
            left, top, right, bottom = win32gui.GetWindowRect(win)  # 取数
            #
            # 最小时点击还原窗口，下面一节为多个窗口，依次点击打开。
            k = 1040  # 最小化后的纵坐标，横坐标约为190
            while top < 0 and k > 800:  # 并设定最多6次，防止死循环
                time.sleep(1)
                win32api.SetCursorPos([180, k - 40])  # 鼠标定位菜单第一个
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                ######点击完成一次
                time.sleep(1)  # 等待窗口出现
                left, top, right, bottom = win32gui.GetWindowRect(win)  # 取数
                if top > 0:  # 判断是否还原
                    break
                else:
                    k -= 40  # 菜单上移一格
                    win32api.SetCursorPos([190, 1040])  # 重新打开菜单
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            win32gui.SetForegroundWindow(win)  # 获取控制
            time.sleep(1)
            return win
        else:
            print('请注意：找不到【%s】这个人（或群），请激活窗口！' % title_name)

    def send_data(self):
        sql = """
        select * from new_option_flow
        where time(original_etl_insert_time ) >= '09:30:00'
        and time(original_etl_insert_time ) <= '17:50:00'
        and weekday(original_etl_insert_time) not in (5,6)
        order by transcation_timestamp asc;
        """

        data = get_data(sql)

        # sql = """
        #         select * from new_option_flow
        #         order by transcation_timestamp asc;
        #         """
        #
        # data = get_data(sql)

        sql_get_history = """
                        select nof.row_hash, replace(his.transcation_time,' ','') ,his.expiry ,his.ContractType ,his.strike,replace(his.ContractSizePrice,' ','') ,his.premium,his.transcation_timestamp 
                        from new_option_flow nof 
                        left join historical_option_flow his 
                        on nof.ticker = his.ticker 
                        and date(nof.transcation_timestamp) = date(his.transcation_timestamp)
                        and time(his.original_etl_insert_time ) between '09:30:00' and '16:29:59';
                        """

        historical_data = get_data(sql_get_history)

        dict_historical_data = dict()

        for row in historical_data:
            if row[0] in dict_historical_data:
                dict_historical_data[row[0]].append(row)
            else:
                dict_historical_data[row[0]] = [row]

        sql_max_ticket = """
        select  ticker,count(*)
        from historical_option_flow s
        where transcation_timestamp between now() - INTERVAL 1 HOUR and now()
        -- where date(transcation_timestamp)='2020-07-20'
        and ticker not in  ('SPY', 'QQQ', 'DIA', 'SPX', 'IWM')
        and time(original_etl_insert_time ) >= '09:30:00'
        and time(original_etl_insert_time ) <= '16:30:00'
        and weekday(original_etl_insert_time) not in (5,6)
        group by ticker
        order by count(*) desc,1
        limit 5;
        """

        sql_max_contract = """
        select ticker , expiry , ContractType ,strike ,count(*)
        from historical_option_flow s
        where transcation_timestamp between now() - INTERVAL 1 HOUR and now()
        -- where date(transcation_timestamp)='2020-07-20'
        and ticker not in  ('SPY', 'QQQ', 'DIA', 'SPX', 'IWM')
        and time(original_etl_insert_time ) >= '09:30:00'
        and time(original_etl_insert_time ) <= '16:30:00'
        and weekday(original_etl_insert_time) not in (5,6)
        group by  ticker , expiry , ContractType ,strike having count(*)>1
        order by count(*) desc,1,2,3,4
        limit 5;
        """

        sql_max_premium = """
        select  ticker  , sum(premium_quantity )
        from historical_option_flow s
        where transcation_timestamp between now() - INTERVAL 1 HOUR and now()
        -- where date(transcation_timestamp)='2020-07-20'
        -- and ticker not in  ('SPY', 'QQQ', 'DIA', 'SPX', 'IWM')
        and time(original_etl_insert_time ) >= '09:30:00'
        and time(original_etl_insert_time ) <= '16:30:00'
        and weekday(original_etl_insert_time) not in (5,6)
        group by ticker
        order by 2 desc,1
        limit 5;
        """

        sql_max_single_premium = """
        select  ticker  , premium_quantity 
        from historical_option_flow s
        where transcation_timestamp between now() - INTERVAL 1 HOUR and now()
        -- where date(transcation_timestamp)='2020-07-20'
        -- and ticker not in  ('SPY', 'QQQ', 'DIA', 'SPX', 'IWM')
        and time(original_etl_insert_time ) >= '09:30:00'
        and time(original_etl_insert_time ) <= '16:30:00'
        and weekday(original_etl_insert_time) not in (5,6)
        order by 2 desc,1
        limit 5;
                """

        max_ticket_data = get_data(sql_max_ticket)

        max_contract_data = get_data(sql_max_contract)

        max_premium_data = get_data(sql_max_premium)

        single_max_premium_data = get_data(sql_max_single_premium)

        conn = build_database_connection()

        cursor = conn.cursor()

        sql_truncate_new = """
                truncate table new_option_flow;
                """

        cursor.execute(sql_truncate_new)

        conn.commit()

        conn.close()

        # free_window = self.local_win(self.free_target)
        # vip_window = self.local_win(self.vip_target)

        global time_point
        print('timpoint:', datetime.datetime.now().hour, time_point)
        # if True:
        if datetime.datetime.now().hour == time_point and 9<=datetime.datetime.now().hour<17:
            print(max_ticket_data)
            print(max_contract_data)
            print(max_premium_data)

            if len(max_contract_data)>0:
                contract_hourly = """最活跃期权 （期权成交数最多）:
"""
                for i in range(len(max_contract_data)):
                    # print(i+1,max_contract_data[i][0],max_contract_data[i][1],max_contract_data[i][2],max_contract_data[i][3])
                    contract = """{}. {} {} {} {}
""".format(i+1,max_contract_data[i][0],max_contract_data[i][1],max_contract_data[i][2],max_contract_data[i][3])
                    contract_hourly += contract
                    print(contract_hourly)
            else:
                contract_hourly = """最活跃期权 （期权成交数最多）:
无重复记录
                                    """





            hourly_update ="""@所有人
最近一小时热点分析
最活跃个股（大单期权数最多）：
1. {}
2. {}
3. {}
4. {}
5. {}

大单成交量（premium）最大：
1. {}   {}
2. {}   {}
3. {}   {}
4. {}   {}
5. {}   {}

单笔大单成交量（premium）最大：
1. {}   {}
2. {}   {}
3. {}   {}
4. {}   {}
5. {}   {}

""".format(max_ticket_data[0][0], max_ticket_data[1][0], max_ticket_data[2][0], max_ticket_data[3][0], max_ticket_data[4][0],
           max_premium_data[0][0], "${:,}".format(max_premium_data[0][1]), max_premium_data[1][0], "${:,}".format(max_premium_data[1][1]),max_premium_data[2][0], "${:,}".format(max_premium_data[2][1]),max_premium_data[3][0], "${:,}".format(max_premium_data[3][1]),max_premium_data[4][0],"${:,}".format(max_premium_data[4][1]),
           single_max_premium_data[0][0],"${:,}".format(single_max_premium_data[0][1]), single_max_premium_data[1][0],"${:,}".format(single_max_premium_data[1][1]), single_max_premium_data[2][0],"${:,}".format(single_max_premium_data[2][1]), single_max_premium_data[3][0],"${:,}".format(single_max_premium_data[3][1]), single_max_premium_data[4][0],"${:,}".format(single_max_premium_data[4][1]))
            for vip_win in self.vip_target:
                vip_window = self.local_win(vip_win)
                self.txt_ctrl_v(hourly_update+contract_hourly)
                self.send_msg(vip_window)
            telegram_bot_sendtext('-1001403437208', hourly_update + contract_hourly)
            # telegram_bot_sendtext('-1001189954337', hourly_update + contract_hourly)
            time_point += 1
        if time_point >= 16:
            time_point = 9

        print(len(data), ' message need to be sent!')
        i = 1

        for row in data:
            line = """{}
Ticker: {} {} {}
Strike: {}

Spot: {}
Size: {}
Type: {}
premium: {}
sector: {}
score: {}
""".format(row[13], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[10], row[11])
            if int(row[11]) >= 90:
                line = """*** High score ***
""" + line
            if 'M' in row[8]:
                line = """$$$ 百万大单 $$$
""" + line
            historical_line = line
            free = ['SPY', 'QQQ', 'DIA', 'IWM']
            if row[12] in dict_historical_data and len(dict_historical_data[row[12]]) > 1:
                historical_records = """
History records:
"""
                for r in sorted(dict_historical_data[row[12]], key=lambda x: x[-1],reverse=True)[:10]:
                    historical_records += r[1] + ' ' + str(r[2]) + ' ' + r[3] + ' ' + str(
                        r[4]) + ' ' + r[5] + ' ' + r[6] + '\n'
                historical_line = line + historical_records

            for vip_win in self.vip_target:
                vip_window = self.local_win(vip_win)

                self.txt_ctrl_v(historical_line)
                self.send_msg(vip_window)

            telegram_bot_sendtext('-1001403437208', historical_line)
            # telegram_bot_sendtext('-1001189954337', historical_line)

            if int(row[11]) >= 90 or 'M' in row[8]:
                for million_win in self.million_target:
                    million_window = self.local_win(million_win)

                    self.txt_ctrl_v(historical_line)
                    self.send_msg(million_window)

            if row[1] in free:
                for f_win in self.free_target:
                    free_window = self.local_win(f_win)

                    self.txt_ctrl_v(line)
                    self.send_msg(free_window)
                # telegram_bot_sendtext('-1001189954337', line)
            print(i, 'messages sent!')
            i += 1
            time.sleep(1)
        print('This session finished!')
        print('')
        print('')
        print('')

    @staticmethod
    def clean_historical_data():
        sql = """
        delete from historical_option_flow 
        where date(original_etl_insert_time) = date(now()) 
        and ((time(original_etl_insert_time ) < '09:29:59' 
        and time(original_etl_insert_time ) > '08:25:00')
        or (time(original_etl_insert_time ) < '16:59:59' 
        and time(original_etl_insert_time ) > '16:14:00') 
        or weekday(original_etl_insert_time) in (5,6));
        """

        conn = build_database_connection()

        cursor = conn.cursor()

        cursor.execute(sql)

        print('Duplicate historical data clean up!')

        conn.commit()

        conn.close()

    def run(self):
        session = self.get_session()
        while True:
            now = datetime.datetime.now()
            if now.isoweekday() in range(1, 6) and now.replace(hour=8, minute=7, second=0, microsecond=0) < now < now.replace(hour=15, minute=25, second=0, microsecond=0):
                global running_count
                print('{} this is {} time run'.format(datetime.datetime.now(), running_count))
                self.extract(session)
                self.remove_duplicates()
                self.send_data()
                if now > now.replace(hour=15, minute=20, second=0, microsecond=0):
                    self.clean_historical_data()
                    print('')
                    print('')
                    print('')
                running_count += 1
                delay = random.randint(3, 6)*15 + random.randint(1, 15)
                print('Next run will start in {}s.'.format(delay))
                print('')
                print('')
                print('')
                time.sleep(delay)

            else:
                break


def _main():
    while True:
        now = datetime.datetime.now()
        if now.isoweekday() in range(1, 6) and now.replace(hour=8, minute=8, second=0, microsecond=0) < now < now.replace(hour=15, minute=20, second=0, microsecond=0):
            try:
                option = Option()
                option.run()
            except:
                telegram_bot_sendtext('-408542611', "Option flow job failed at {}".format(str(datetime.datetime.now())))
                time.sleep(120)
        else:
            delay = ((datetime.timedelta(hours=24) - ( now.replace(second=0, microsecond=0) - now.replace(hour=8, minute=10, second=0, microsecond=0))).total_seconds() % ( 24 * 3600))
            if delay == 0:
                delay = 60
            print('Out of transcation window, current time is {}, waiting time is {}s'.format(now, delay))
            time.sleep(delay)


if __name__ == '__main__':
    _main()
