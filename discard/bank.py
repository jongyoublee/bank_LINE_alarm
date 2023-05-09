from simple_bank_korea.kb import get_transactions

transaction_list = get_transactions(
    bank_num = '85673704010002',    # 계좌번호
    birthday = '1098211678',        # 사업자번호
    password = '1350',              # 계좌비밀번호
    days = 3,                       # 조회할 기간(기본 30일)
    PHANTOM_PATH = './phantomjs-2.1.1-windows/bin/phantomjs.exe',
)

for trs in transaction_list:
    if(trs['amount'] > 0):
        print(trs['date'], format(trs['amount'], ','), trs['transaction_by'])
    else:
        continue