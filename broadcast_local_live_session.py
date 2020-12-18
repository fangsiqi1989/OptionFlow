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

# python C:\Users\Siqi\Desktop\OptionFlow\broadcast_local_live_session.py -url https://app.flowalgo.com/ -login_url https://app.flowalgo.com/users/login -proxy http:5.79.66.2:13010 -username option20201001 -password option123 -free_target "Option Flow Fre" -vip_target "Option Flow VI"

parser = argparse.ArgumentParser()
parser.add_argument('-url', type=str, default=None)
parser.add_argument('-login_url', type=str, default=None)
parser.add_argument('-proxy', type=str, default=None)
parser.add_argument('-username', type=str, default=None)
parser.add_argument('-password', type=str, default=None)
parser.add_argument('-free_target', type=str, default=None)
parser.add_argument('-vip_target', type=str, default=None)
args = parser.parse_args()


headers = [
{
'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
,'cookie': '__cfduid=df32cb19205bc7e4b8ba14a3ccd5d96cf1601614085; _ga=GA1.2.1122056143.1601614090; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; intercom-id-dtoll8e6=5f313d1c-1f95-4882-8841-ca87260d4020; _fbp=fb.1.1601614094994.719174546; __adroll_fpc=ce481f124e8b1fd7a3839ad834766b81-1601960955579; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201005%3A4%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201005%3A4%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201005%3A4; _gid=GA1.2.1458854780.1603428499; _gat_gtag_UA_105239038_2=1; PHPSESSID=26e6ba6a8d9122d1a96b5cb9eb4d184f; amember_nr=5b0a78a9b68f388ea3ffb0014cc2e7b2; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1603681991%7C4Rqd972djksGLT73DgrFE6Mgt4McSWq8oMWZLvtgzCG%7Cfdecdf1f8e9c7d5289ad214d71b441978da77a1cd3c2a365d474e915c6f51ddf; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2015595%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%7D; intercom-session-dtoll8e6=SXdoMWNZQjFsR3VkZVZrSE50aHgrbkdDbmZFUzEvNlQ5VXNYU2I1VzF2QkhSdEtvMjlEL3VzVkVCaU50WWlrKy0tQkczdk9tZWYxNHI3UnVMa2FONkl6UT09--ae3d4db7825b214e14e71ddd949cc4891bcfd91b; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201005%3A5%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201005%3A4%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201005%3A4'
},
{
'user-agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36'
,'cookie':'__cfduid=df32cb19205bc7e4b8ba14a3ccd5d96cf1601614085; _ga=GA1.2.1122056143.1601614090; SL_C_23361dd035530_VID=xSbHHwMDqY-; SL_C_23361dd035530_KEY=909c690836dff219fdb765f6e1091e5a99e5f112; __adroll_fpc=282d4f361d234d5aa246286eee869388-1601614094159; intercom-id-dtoll8e6=5f313d1c-1f95-4882-8841-ca87260d4020; _fbp=fb.1.1601614094994.719174546; __adroll_fpc=ce481f124e8b1fd7a3839ad834766b81-1601960955579; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201005%3A4%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201005%3A4%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201005%3A4; amember_nr=b77a126af22d23407e7f365f902dda6e; _gid=GA1.2.1458854780.1603428499; _gat_gtag_UA_105239038_2=1; PHPSESSID=80a6a6cc02e2315510e15e13c49280a1; wordpress_logged_in_d1f53b3265d55ab79282aac86fcd5ba4=option20201001%7C1603601307%7CxusB3XtTtzRt40EwSMAMMFH1DRSREUNSuEKpK7yfnAq%7C2cd002cc61eea92826e0df068216b5c0af829b61c17e2f2b2e588c0d1cd3e874; mp_cef79b4c5c48fb3ec3efe8059605ec56_mixpanel=%7B%22distinct_id%22%3A%2015595%2C%22%24device_id%22%3A%20%22174e7a397b8243-07855eaa3ee8d-333376b-e1000-174e7a397b93d9%22%2C%22%24initial_referrer%22%3A%20%22https%3A%2F%2Fflowalgo.com%2F%22%2C%22%24initial_referring_domain%22%3A%20%22flowalgo.com%22%2C%22%24user_id%22%3A%2015595%2C%22__mps%22%3A%20%7B%7D%2C%22__mpso%22%3A%20%7B%7D%2C%22__mpus%22%3A%20%7B%7D%2C%22__mpa%22%3A%20%7B%7D%2C%22__mpu%22%3A%20%7B%7D%2C%22__mpr%22%3A%20%5B%5D%2C%22__mpap%22%3A%20%5B%5D%7D; intercom-session-dtoll8e6=bmd2T2NxbXFHdWxodVg2K1hELzhlNWpHeTdZR09sZ0tEY2VWUkNnMk1Kd1R6dWJUTTd1ckpXWW1WT3h0NE1JdS0tR2RpV0hpd3REMHl6ZnFpSDRWZ3R4Zz09--5c6fcd2e1cb4407f793daa174b971b625993415f; __ar_v4=NM6BA5VRZNDFJDSOSRXRPP%3A20201005%3A5%7CJURQBX5ZWNGNLAFB4ISSVP%3A20201005%3A4%7CEWEJP57Y6NEUVKJVXOCEJP%3A20201005%3A4'
}

]
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
            num = random.randint(0, len(headers) - 1)
            response = session.get(self.url, headers=headers[datetime.datetime.now().day%2])
            # response = session.get(self.url, proxies=proxy)

            parser = etree.HTMLParser(encoding='utf-8')

            htmlElement = etree.fromstring(response.content.decode('utf-8'), parser=parser)

            data = htmlElement.xpath(
                '//*[@id="optionflow"]/div[2]//div[@class and @data-ticker and @data-sentiment and @data-flowid and @data-premiumpaid and @data-ordertype]')
            print('=' * 100)
            print('=' * 100)
            print('=' * 100)
            print('=' * 100)
            # print(response.content.decode('utf-8'))
            print(datetime.datetime.now())
            print('Number of records', len(data))
            if len(data) == 0:
                self.telegram_bot_sendtext('-408542611', 'No records return from website!!!')
            now = datetime.datetime.now()
            if now.replace(hour=8, minute=0, second=0, microsecond=0) < now < now.replace(hour=15, minute=30, second=0, microsecond=0):
                delay = random.randint(3, 6) * 15 + random.randint(1, 15)
                print('Next run will start in {}s.'.format(delay))
                time.sleep(delay)
            else:
                break




def _main():
    global running_count
    while True:
        now = datetime.datetime.now()
        if now.replace(hour=8, minute=0, second=0, microsecond=0) < now < now.replace(hour=15, minute=30, second=0, microsecond=0):
            print('{} this is {} time run'.format(datetime.datetime.now(),running_count))
            extract = Extract()
            extract.run()
        else:
            delay = ((datetime.timedelta(hours=24) - (now.replace(second=0, microsecond=0) - now.replace(hour=8, minute=10, second=0, microsecond=0))).total_seconds() % (24 * 3600))
            if delay == 0:
                delay = 60
            print('Out of transcation window, current time is {}, waiting time is {}s'.format(now, delay))
            time.sleep(delay)



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

