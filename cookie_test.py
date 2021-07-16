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

# python C:\Users\Siqi\Desktop\OptionFlow\cookie_test.py -url https://app.flowalgo.com/ -login_url https://app.flowalgo.com/users/login -proxy http:5.79.66.2:13010 -username option20210115 -password option123 -free_target "Option Flow Fre" -vip_target "Option Flow VI"

parser = argparse.ArgumentParser()
parser.add_argument('-url', type=str, default=None)
parser.add_argument('-login_url', type=str, default=None)
parser.add_argument('-proxy', type=str, default=None)
parser.add_argument('-username', type=str, default=None)
parser.add_argument('-password', type=str, default=None)
parser.add_argument('-free_target', type=str, default=None)
parser.add_argument('-vip_target', type=str, default=None)
args = parser.parse_args()


headers ={
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
,'cookie':'_ga=GA1.2.1122056143.1601614090; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; _fbp=fb.1.1601614094994.719174546; __adroll_fpc=ce481f124e8b1fd7a3839ad834766b81-1601960955579; _gid=GA1.2.1343756811.1626328784; _gat_gtag_UA_105239038_2=1; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP:20210713:1|JURQBX5ZWNGNLAFB4ISSVP:20210713:1|EWEJP57Y6NEUVKJVXOCEJP:20210713:1; PHPSESSID=084a5232a09520ee5ffada6edee4b89c; amember_nr=a5234b01dc59b6fca2d812761f404063; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20210706|1626501589|wmmKS7tSM16NA8RC27TuaykxYezCDgqAUoSRfHw9Xud|fe9432aad2faf7a15fb9ba071a65c5e5deeedc3469a339eb8b870caa99bb0fb8; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel={"distinct_id": 27495,"$device_id": "174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9","$initial_referrer": "https://flowalgo.com/","$initial_referring_domain": "flowalgo.com","$user_id": 27495,"__mps": {},"__mpso": {},"__mpus": {},"__mpa": {},"__mpu": {},"__mpr": [],"__mpap": []}; intercom-session-dtoll8e6=RzBtSTJSMFhkYkpsM0gwZjZDdWZwWkpOaXdZNTJIRWpDVEVzcHpUYW0rUW5VNkNZc0pQekhQamZFRk9lYkhCRy0tNllYYTA4WlJrUUFyK2p2S2tWaytRZz09--e6d8c7d415f3644895840145d4b04c367f6d0c01; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP:20210713:2|JURQBX5ZWNGNLAFB4ISSVP:20210713:1|EWEJP57Y6NEUVKJVXOCEJP:20210713:1'

}


running_count = 1


class Extract:
    def __init__(self):
        self.url = args.url
        self.login_url = args.login_url
        self.proxy = {args.proxy.split(':')[0]: args.proxy.split(':')[1] + ':' + args.proxy.split(':')[2]}
        self.username = args.username
        self.password = args.password

    def telegram_bot_sendtext(chatID, bot_message):
        print('execute telegram_bot_sendtext')
        bot_token = '1323919359:AAEtt77oSWn4rHxExKvSB3QDmyZu9hnGblM'
        bot_chatID = chatID
        send_text = 'https://api.telegram.org/bot' + bot_token + '/sendMessage?chat_id=' + bot_chatID + '&parse_mode=Markdown&text=' + bot_message

        response = requests.get(send_text)

        return response.json()

    def run(self):
        data = {
            'amember_login': self.username
            , 'amember_pass': self.password
        }

        session = requests.session()

        session.post(self.login_url, data)
        while True:
            now = datetime.datetime.now()
            response = session.get(self.url, headers=headers, timeout=10)

            parser = etree.HTMLParser(encoding='utf-8')

            htmlElement = etree.fromstring(response.content.decode('utf-8'), parser=parser)

            data = htmlElement.xpath(
                '//*[@id="optionflow"]/div[2]//div[@class and @data-ticker and @data-sentiment and @data-flowid and @data-premiumpaid and @data-ordertype]')
            # print(response.content.decode('utf-8'))
            if len(data)>0:
                print(datetime.datetime.now())
                print('Number of records', len(data))
            else:
                print('='*30)
                print('No data return!')
                print('='*30)

            delay = random.randint(3, 6) * 15 + random.randint(1, 15)
            print('Next run will start in {}s.'.format(delay))
            print('')
            print('')
            print('')
            time.sleep(delay)





def _main():
    global running_count

    print('{} this is {} time run'.format(datetime.datetime.now(),running_count))
    extract = Extract()
    extract.run()




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

