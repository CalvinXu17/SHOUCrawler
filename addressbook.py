import requests
from bs4 import BeautifulSoup
import xlwt
import time

header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3676.400 QQBrowser/10.4.3469.400',
}

url1 = 'https://uis.shou.edu.cn/cas/login?isLoginService=11&service=http://ecampus.shou.edu.cn/c/portal/login'
urlf = 'http://ecampus.shou.edu.cn/web/guest/addressbook?p_p_id=shouAddressList_WAR_shouAddressListportlet&p_p_lifecycle=0&p_p_state=exclusive&p_p_mode=view&p_p_col_id=column-1&p_p_col_count=1&_shouAddressList_WAR_shouAddressListportlet_action=load&_shouAddressList_WAR_shouAddressListportlet_cur='
urls = '&_shouAddressList_WAR_shouAddressListportlet_delta=20'

usern = 'yourid'
passw = 'yourpassword'
datas = {'username': usern, 'password': passw, 'submit': ''}


def GET_COOKIE():
    rq = requests.session()
    rq.post(url1, data=datas, headers=header)
    return rq


def GET(url, rq):
    try:
        ret = rq.get(url, headers=header)
        print(ret.status_code)
        ret.encoding = 'utf-8'
        return ret.text
    except:
        print('Wrong')
        return ''


def start(rq):
    judge = True
    sheet = []

    for i in range(95):
        t = i + 1
        ret = GET(urlf + str(t) + urls, rq)
        bs = BeautifulSoup(ret, 'html.parser')

        title = bs.tr.find_all('th')
        tt = []

        if judge:
            for m in title:
                tt.append(m.string)
            judge = False

        # print(tt)
        if (len(tt) > 0):
            sheet.append(tt)

        for tr in bs.find('tr').next_siblings:
            ll = []
            for td in tr:
                try:
                    for p in td:
                        if p.string == None:
                            ll.append('')
                        elif p.string != '\n':
                            ll.append(p.string)
                except:
                    pass
            if len(ll) > 0:
                sheet.append(ll)

        time.sleep(0.1)

    return sheet


def write(sheet):
    i = 0
    j = 0

    workbook = xlwt.Workbook(encoding='utf-8')
    sheet1 = workbook.add_sheet('address')

    try:
        for x in sheet:
            j = 0
            for y in x:
                sheet1.write(i, j, y)
                j += 1
            i += 1
        workbook.save('D://adr.xls')
    except:
        print('Wrong')
        pass


if __name__ == '__main__':

    rq = GET_COOKIE()
    sheet = start(rq)

    for i in sheet:
        print(i)

    print('total:' + str(len(sheet)))

    write(sheet)
