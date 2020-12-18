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

# python C:\Users\Siqi\Desktop\OptionFlow\broadcast_local_proxy.py -url https://app.flowalgo.com/ -login_url https://app.flowalgo.com/users/login -proxy http:5.79.66.2:13010 -username option20201001 -password option123 -free_target "Option Flow Fre" -vip_target "Option Flow VI"


parser = argparse.ArgumentParser()
parser.add_argument('-url', type=str, default=None)
parser.add_argument('-login_url', type=str, default=None)
parser.add_argument('-proxy', type=str, default=None)
parser.add_argument('-username', type=str, default=None)
parser.add_argument('-password', type=str, default=None)
parser.add_argument('-free_target', type=str, default=None)
parser.add_argument('-vip_target', type=str, default=None)
args = parser.parse_args()


headers = [{
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:81.0) Gecko/20100101 Firefox/81.0'
,'cookie':'__adroll_fpc=ec138dc5651bbffdf6809a94abc7dd73-1601564534662; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200931%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200931%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20200931%3A2; _fbp=fb.1.1601563852978.15442719; __cfduid=d56cb17662ac48aad2bcab65b8a2bd8941601563853; PHPSESSID=9e9ec6eb99e4ef5c48dc9e54b858e5c6; _ga=GA1.2.1217111726.1601564532; _gid=GA1.2.606489091.1601564532; intercom-id-dtoll8e6=22630759-c915-4933-a8e0-bf4b7b302598; intercom-session-dtoll8e6=SHNWcm5YN0lYMEx1cGNmNGcxT1ppZUpmRzJNWG…efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174e4b108122be-01a1fbc4892f188-4c3f257b-e1000-174e4b10813203%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2Fselect-a-plan%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%2C%22%24user_id%22%3A%2015595%7D'.encode('utf-8')
},
{
# mac pro firefox
'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.14; rv:81.0) Gecko/20100101 Firefox/81.0'
,'cookie':'_ga=GA1.2.963995471.1590266285; intercom-id-dtoll8e6=765bde21-ca98-4aec-81f4-d8808fc5db3a; __adroll_fpc=115ce5abb5ab7585140ac4104ac03857-1590266307580; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200912%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200912%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20200912%3A2; _fbp=fb.1.1590266289963.1401559072; __adroll_fpc=115ce5abb5ab7585140ac4104ac03857-1590266307580; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200912%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200912%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20200912%3A1; SL_…_UA_105239038_2=1; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%221724341cf5522c-0e7337f5b8c0408-4b5569-1fa400-1724341cf5680c%22%2C%22%24initial_referrer%22%3A%20%22%24direct%22%2C%22%24initial_referring_domain%22%3A%20%22%24direct%22%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%2C%22%24user_id%22%3A%2015595%7D'.encode('utf-8')
},
{
# mac pro chrome
'user-agent':'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36'
,'cookie':'__cfduid=dceb769e84c3ba2adde48ff6a05095db21600889952; PHPSESSID=af21240b2a370bfb6b52b91d5b56191b; _ga=GA1.2.1693030311.1600957852; __adroll_fpc=2e5bcdc5d8344fe06493fd29db21ea16-1600957853471; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; _fbp=fb.1.1600957853843.836245199; intercom-id-dtoll8e6=3c88aacb-ea2e-4e9d-a899-4a2287ee1db0; _gid=GA1.2.2092655903.1601613287; _gat_gtag_UA_105239038_2=1; SL_C_23361dd035530_VID=u3vxmiOHvE; SL_C_23361dd035530_SID=EUDhh8KvbB; amember_nr=b7aaa37c69ab13d059984ad11229bb42; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1601786100%7COp3FnKniVA3G95sLrl8UdYF7XqlQw2wPfrEUccSTjxT%7Cc9fbc2dcf7b06f1946d2c559c68555d1c19c85c05b7def595827a8c4641a8387; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174c0862aa7aa6-011cbff22c1c99-316b7005-13c680-174c0862aa88e6%22%2C%22%24initial_referrer%22%3A%20%22%24direct%22%2C%22%24initial_referring_domain%22%3A%20%22%24direct%22%2C%22%24user_id%22%3A%2015595%7D; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20200924%3A16%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200924%3A16%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200924%3A16; intercom-session-dtoll8e6=eDh5UFN3ZFByckVlUGJFNTk3Q3NzVGlpeFFQb3o0bFN2b0hMZmhWN3JHVUZIaTR3bzBsTHVBTW5MQVBqUkFBby0tNm8zdE5mSDBhSWJWSU8yVnNsWThuQT09--1a2f60a6b799672c8048634f9700498f3a91ec2d'
},
{
# dell edge
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18362'
,'cookie':'__adroll_fpc=b31410d05e81487e93800fba7406c094-1601261941695; __ar_v4=EWEJP57Y6NEUVKJVXOCEJP%3A20200928%3A5%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200928%3A5%7CNM6BA5VRZNDFJDSOSRXRPP%3A20200928%3A6; SL_C_23361dd035530_SID=x8JuLuA6uAb; __adroll_fpc=b31410d05e81487e93800fba7406c094-1601261941695; SL_C_23361dd035530_VID=FY9P6_1at-M; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200928%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200928%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20200928%3A1; amember_nr=dcc7022b3d89ec6b3f1f6b7a590e42b3; _gid=GA1.2.429820660.1601562557; intercom-id-dtoll8e6=1e0f5910-abe5-406b-a517-778fe8cda0f4; intercom-session-dtoll8e6=TGVDMVE3VTZHdHBJaXpsZmMrMGlFUHdKMTNiaHJOVWxEK0FSc2VZOEZKSGRQOURlanlVcm1RSC8zTTgrcGh6ci0ta2JOcFowaFRueDFCbk80QWYxZ1o1dz09--0d585cef8ddd6775c2d4417d79da90e12b0e343a; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; _ga=GA1.2.284200782.1601261939; __cfduid=db12bbeb205a6b81e472e331665b1353d1600871497; _fbp=fb.1.1601261943782.1481715877; PHPSESSID=bcef61bba7a7cd98f8474d54d71aeaa8; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1601786669%7CbIMAwX25SsRLoP8fOPSF7kLuaMUeYACOlph65ENTZCS%7C7a7f24067a4587c511f5394a6a337347a11483d60da4aff11e34cdbe3782ee06; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174d2a6539be6-0c85953b091a7-71415a3b-e1000-174d2a6539c1c5%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%2C%22%24user_id%22%3A%2015595%2C%22%24search_engine%22%3A%20%22bing%22%2C%22mp_keyword%22%3A%20%22flowalgo%22%7D'
},
{
# dell chrome
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
,'cookie':'__cfduid=df32cb19205bc7e4b8ba14a3ccd5d96cf1601614085; _ga=GA1.2.1122056143.1601614090; _gid=GA1.2.200118318.1601614090; _gat_gtag_UA_105239038_2=1; PHPSESSID=31b465cf6ed8eb3a2095326f74cbfd4c; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; intercom-id-dtoll8e6=5f313d1c-1f95-4882-8841-ca87260d4020; SL_C_23361dd035530_SID=N8pmVB-0K7J; _fbp=fb.1.1601614094994.719174546; amember_nr=083b117d309e7e8190beebb53db959f2; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1601786904%7Cukmp0AgC2RmzN71VRXUZZhls05toGaf64yOmFzePAVc%7C71efc2558a5bf9c678dd74688a808b1feca76e68b1e553c9710a1d5f551b36a7; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2015595%7D; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201001%3A2%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201001%3A2%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201001%3A2; intercom-session-dtoll8e6=MVFkYTNLNFVXd2QwcjRXaER2NFJkbHJKNTZIU3pOdUR6Z21aU09LcFV5RFpSKzR3Q2lXdWRnNnRwRlRIR08xVS0tWTRwNGJtTUErQkwvbVo5NHJNWEtJQT09--07265877c5e478f82a8e333caccf7e1e79271cb2'
},
{
# dell firefox
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:81.0) Gecko/20100101 Firefox/81.0'
,'cookie':'__cfduid=d9575aeed23abc90e2ae7310bf13769c61601614211; _ga=GA1.2.882484519.1601614214; _gid=GA1.2.1342851212.1601614214; intercom-id-dtoll8e6=1b3cec80-3400-405e-b8e8-aa6ac187b240; intercom-session-dtoll8e6=; __adroll_fpc=52aae2a98e3d57c84baee7508a8c5dae-1601614217675; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201001%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201001%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20201001%3A1; _fbp=fb.1.1601614219456.357098368; PHPSESSID=3f400d17e4c782213a8efa252390da6e; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%20%22174e7a664067e-0c2f9365bbf568-4c3f257b-e1000-174e7a664078c%22%2C%22%24device_id%22%3A%20%22174e7a664067e-0c2f9365bbf568-4c3f257b-e1000-174e7a664078c%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%7D; _gat_gtag_UA_105239038_2=1; SL_C_23361dd035530_VID=0bgDW4PUV; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; SL_C_23361dd035530_SID=jXzxaXOR-f5; __adroll_fpc=52aae2a98e3d57c84baee7508a8c5dae-1601614217675; __ar_v4=%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201001%3A1%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201001%3A1%7CNM6BA5VRZNDFJDSOSRXRPP%3A20201001%3A2; amember_nr=209cc92b2bd2120412e1d6f23cddee6a; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1601787091%7CQFG9oPUNcMUF2e1v9cJy3xaY1u0tQikiEFHWhkcPMyF%7Cfb3ae613951b9f63c6753513ecb4054a7d4dfed7e3a42bad974557fb311aebab'.encode('utf-8')
}
]
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


class Extract:
    def __init__(self):
        self.url = args.url
        self.login_url = args.login_url
        self.proxy = {args.proxy.split(':')[0]: args.proxy.split(':')[1] + ':' + args.proxy.split(':')[2]}
        self.username = args.username
        self.password = args.password

    def get_session(self):
        data = {
            'amember_login': self.username
            , 'amember_pass': self.password
        }

        session = requests.session()
        session.post(self.login_url, data, headers)
        print('Session created!')
        return session


    def extract(self, session):
        proxy = {
            # 'http':'http://206.189.235.214:80',
            # 'https':'https://206.189.235.214:80'
            # world wide
            "http": "http://5.79.73.131:13010",
            "https": "https://5.79.73.131:13010"
            # USA
            # "http": "http://5.79.66.2:13010",
            # "https": "https://5.79.66.2:13010"
        }

        # print(proxy)
        # print(session)
        num = random.randint(0,len(headers)-1)
        # print(list_cookies)
        # headers = {
        #     'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18362'
        #
        #     # ,'cookie': list_cookies[num]
        #     , 'cookie':'__adroll_fpc=8650dd071d4f9ee9d10afdfe08713b16-1598885541788; intercom-id-dtoll8e6=7527060d-1790-4fb5-af0a-943ba3479f97; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20200830%3A2%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200830%3A2%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200830%3A2; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; _ga=GA1.2.1459804922.1598885544; _gac_UA-105239038-2=1.1598885544.EAIaIQobChMIwd374NjF6wIVe-bjBx3SFQxoEAAYASAAEgLeG_D_BwE; __adroll_fpc=8650dd071d4f9ee9d10afdfe08713b16-1598885541788; _fbp=fb.1.1598885542545.864320631; __cfduid=d963325496004b44dfb232b031f2e0b451601347013; intercom-session-dtoll8e6=WVpJNzhjRUlRcGRuNnU3M0dyV2FWK1AyN2pmTVI0S3N6eUZwUlV2M0UrOEIxeWcvU3ZKV1FKWncweEJBRDFDcC0tY2x5SkVwTEt1ZEN4TFNYQnhEYlV2Zz09--f4c4eb9742eb730746de6486e0da12395797ed2d; _gid=GA1.2.1645237115.1601049796; _gat_gtag_UA_105239038_2=1; PHPSESSID=69fd31b518f6e2ded7cbc00bf0d7948c; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015254%2C%22%24device_id%22%3A%20%2217445014dee19b-0c64897018d43f-8383667-e1000-17445014def1fe%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%3Fgclid%3DEAIaIQobChMIwd374NjF6wIVe-bjBx3SFQxoEAAYASAAEgLeG_D_BwE%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2015254%7D; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20200830%3A3%7CJURQBX5ZWNGNLAFB4ISSVP%3A20200830%3A2%7CEWEJP57Y6NEUVKJVXOCEJP%3A20200830%3A2; amember_nr=2dbc151f8b7fc1d87fbb368cde25c975; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=audreyb123%7C1601222601%7CERHt2edwbNvzkklG9SaVw8yUHW11M5xN7YuoMRQ4w8f%7Cee60a488af286c94613d2552e07d7e29db84a284ff62056cdab01733b177589b'
        # }
        # response = session.get(self.url, proxies=proxy)
        response = session.get(self.url, headers=headers[num], proxies=proxy)
        # response = session.get(self.url, headers=headers)

        # response = session.get(self.url, headers=headers, proxies=proxy)
        # response = session.get(self.url, headers=headers)
        # print(response.content.decode('utf-8'))
        # with open('flowalgo.html', 'w') as fp:
        #     fp.write(response.content.decode('utf-8'))

        parser = etree.HTMLParser(encoding='utf-8')
        # print(parser)

        htmlElement = etree.fromstring(response.content.decode('utf-8'), parser=parser)
        # htmlElement = etree.parse('flowalgo.html', parser=parser)
        # print(htmlElement)

        data = htmlElement.xpath('//*[@id="optionflow"]/div[2]//div[@class and @data-ticker and @data-sentiment and @data-flowid and @data-premiumpaid and @data-ordertype]')
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        print('=' * 100)
        # print(response.content.decode('utf-8'))
        print('Number of records', len(data))
        # if len(data) == 0:
        #     print(response.content.decode('utf-8'))
        #     telegram_bot_sendtext('-408542611', 'No records return from website!!!')
        sys.exit(0)
        conn = build_database_connection()



        cursor = conn.cursor()

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
        TRIM(TRAILING 'K' FROM TRIM(LEADING '$' from rt.premium))*1000,
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

    def run(self):
        session = self.get_session()
        self.extract(session)
        self.remove_duplicates()


class Broadcast:
    def __init__(self):
        self.free_target = args.free_target.split(',')
        self.vip_target = args.vip_target.split(',')

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
        and time(original_etl_insert_time ) <= '16:50:00'
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

            if row[1] in free:
                for f_win in self.free_target:
                    free_window = self.local_win(f_win)

                    self.txt_ctrl_v(line)
                    self.send_msg(free_window)
                telegram_bot_sendtext('-1001189954337', line)
            print(i, 'messages sent!')
            i += 1
            time.sleep(1)
        print('This session finished!')
        print('')
        print('')
        print('')

    def clean_historical_data(self):
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


def _main():
    global running_count
    print('{} this is {} time run'.format(datetime.datetime.now(),running_count))
    extract = Extract()
    extract.run()
    # broadcast = Broadcast()
    # broadcast.send_data()
    # running_count += 1
    # now = datetime.datetime.now()
    # if now > now.replace(hour=15, minute=16, second=0, microsecond=0):
    #     broadcast.clean_historical_data()
    #     print('')
    #     print('')
    #     print('')

    # try:
    #     global running_count
    #     print('{} this is {} time run'.format(datetime.datetime.now(),running_count))
    #     extract = Extract()
    #     extract.run()
    #     # broadcast = Broadcast()
    #     # broadcast.send_data()
    #     # running_count += 1
    #     # now = datetime.datetime.now()
    #     # if now > now.replace(hour=15, minute=16, second=0, microsecond=0):
    #     #     broadcast.clean_historical_data()
    #     #     print('')
    #     #     print('')
    #     #     print('')
    # except:
    #     telegram_bot_sendtext('-408542611', "Option flow job failed at {}".format(str(datetime.datetime.now())))
    #     time.sleep(240)


if __name__ == '__main__':
    _main()
    # while True:
    #     now = datetime.datetime.now()
    #     if now.isoweekday() in range(1, 6) and now.replace(hour=8, minute=15, second=0,
    #                                                        microsecond=0) < now < now.replace(hour=15, minute=20,
    #                                                                                           second=0, microsecond=0):
    #         try:
    #             _main()
    #             time.sleep(240)
    #         except:
    #             telegram_bot_sendtext('-408542611', "Option flow job failed at {}".format(str(datetime.datetime.now())))
    #             time.sleep(240)
    #     else:
    #         delay = ((datetime.timedelta(hours=24) - (
    #                     now.replace(second=0, microsecond=0) - now.replace(hour=8, minute=20, second=0,
    #                                                                        microsecond=0))).total_seconds() % (
    #                              24 * 3600))
    #         if delay == 0:
    #             delay = 60
    #         print(max(10, datetime.datetime.now().hour+1 if datetime.datetime.now().hour+1<16 else 10))
    #         print('Out of transcation window, current time is {}, waiting time is {}s'.format(now, delay))
    #         time.sleep(delay)





    # scheduler = BlockingScheduler()
    # scheduler.add_job(func=_main, trigger='interval', minutes=3)
    # scheduler.start()

