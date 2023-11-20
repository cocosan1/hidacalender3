import streamlit as st
import pandas as pd
import openpyxl
import datetime
from dateutil.relativedelta import relativedelta

import urllib.request
import tabula
from io import BytesIO

st.set_page_config(page_title='納期カレンダー作成')
st.markdown('#### 納期カレンダー作成')

public_holiday_csv_url="https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
public_holiday2 = public_holiday_csv_url.split("/")[-1] #ファイル名の取り出し

# 対象年度
this_year = datetime.date.today().year
next_year = (datetime.date.today() + relativedelta(years=1)).year

# 会社特有の休日
company_holiday_shukka = [ 

    '2023/12/26', '2023/12/27', '2023/12/28',\
    '2023/12/29', '2023/12/30', '2023/12/31',\
    '2024/1/1', '2024/1/2', '2024/1/3', '2024/1/4', '2024/1/5', '2024/1/6',
    '2024/4/27', '2024/4/30', '2024/5/1', '2024/5/2',
    '2024/8/10', '2024/8/13', '2024/8/14', '2024/8/15', '2024/8/16', '2024/8/17'
    ]

with st.expander('会社の休日設定', expanded=False):
    st.write(company_holiday_shukka)

company_holiday_chakubi = [

    '2023/12/26', '2023/12/27', '2023/12/28',\
    '2023/12/29', '2023/12/30', '2023/12/31',\
    '2024/1/1', '2024/1/2', '2024/1/3', '2024/1/4', '2024/1/5', '2024/1/6',
    '2024/4/27', '2024/4/30', '2024/5/1', '2024/5/2',
    '2024/8/10', '2024/8/13', '2024/8/14', '2024/8/15', '2024/8/16', '2024/8/17'
    ]

kadoubi_this = []
chakubi_this = []
kadoubi_next = []
chakubi_next = []

option_day = 0

# ***ファイルアップロード 今期***
uploaded_file = st.file_uploader('出荷日表PDFの読み込み', type='pdf', key='shukka')
df = pd.DataFrame()
if not uploaded_file:
    st.info('出荷日表PDFを選択してください。')
    st.stop() 
elif uploaded_file:
    df = tabula.read_pdf(uploaded_file, lattice=True, pages='1') #dfのリストで出力される

###### 出荷日の次の日から何日後を着日とするか
day_list = [1, 2, 3, 4, 5, 6, 7]
option_day = st.radio(
    "出荷日の次の日から何日後を着日とするか？（稼働日）",
    day_list, index=4
)

#稼働日判定/出荷・移動日
def get_kadoubi(date):
    # 0埋め解消　祝日ファイルに合わせて
    year = date.strftime("%Y")
    month = date.strftime("%m").lstrip("0") #strの左から0を削除
    day = date.strftime("%d").lstrip("0")

    # 日曜日
    if (date.weekday() == 6): # 0月曜 6日曜
        return False
    
    # 祝日
    holidays_df = pd.read_csv(public_holiday2, encoding="SHIFT-JIS")
    if date.strftime(year + "/" + month + "/" + day) in holidays_df['国民の祝日・休日月日'].tolist():
        return False 

    #　会社の休日
    if date.strftime(year + "/" + month + "/" + day) in company_holiday_shukka:
        return False

    return True 


#着日判定
def get_chakubi(date):
    # 0埋め解消　祝日ファイルに合わせて
    year = date.strftime("%Y")
    month = date.strftime("%m").lstrip("0")
    day = date.strftime("%d").lstrip("0")

    # 日曜日
    if (date.weekday() == 6): # 0月曜 6日曜
        return False
    
    # 水曜
    if (date.weekday() == 2): # 0月曜 6日曜
        return False    

    # 祝日
    holidays_df = pd.read_csv(public_holiday2, encoding="SHIFT-JIS")
    if date.strftime(year + "/" + month + "/" + day) in holidays_df['国民の祝日・休日月日'].tolist():
        return False

    #　会社の休日
    if date.strftime(year + "/" + month + "/" + day) in company_holiday_chakubi:
        return False

    return True

def generate_pdf():

    cols = ['SEOTO-EX', 'Aパターン', 'Bパターン', '30日']
    cols_nonex = ['Aパターン', 'Bパターン', '30日']
    
    #表が格子状になっている場合 lattice=True そうでない　stream=True　複数ページ読み込み pages='all'
    df_calend = df[0]
    df_calend = df_calend.dropna(how='any')
    df_calend = df_calend.drop(df_calend.columns[[5, 6, 7]], axis=1) #40日から右カラムの削除
    df_calend = df_calend.rename(columns={'Unnamed: 0': '受注日', 'KX250AX\rKX260AX': 'SEOTO-EX'})

    #曜日を消す
    for col in cols_nonex:
        df_calend[col] = df_calend[col].str[:-2]
    
    #2022年の追加
    for col in cols:
        df_calend[col] = f'{this_year}年' + df_calend[col]

    #datetime型に変換
    for col in cols:
        df_calend[col] = pd.to_datetime(df_calend[col], format='%Y年%m月%d日')

    #年またぎ対応
    #today datetime64型の用意　＊一度dfにしてから変換
    today = datetime.date.today()
    temp_list = []
    temp_list.append(today)
    df_temp = pd.DataFrame(temp_list, columns=['date'])
    df_temp['date2'] = pd.to_datetime(df_temp['date'], format='%Y-%m-%d')
    today2 = df_temp['date2'][0]

    #todayより前のものは年を来年に変換
    for col in cols:
        df_calend[col] = df_calend[col].map(lambda x: x + relativedelta(years=1) if x < today2 else x)

    #　str型へ　時間を消す
    # Series型にはdtというアクセサが提供されており、Timestamp型を含む日時型の要素から、
    # 日付や時刻のみの要素へ一括変換できる。
    for col in cols:
        df_calend[col] = df_calend[col].dt.strftime('%Y-%m-%d')

    #着日計算
    arrival_ex = []
    arrival_a = []
    arrival_b = []
    arrival_30 = []

    arrivals = [arrival_ex, arrival_a, arrival_b, arrival_30]

    i = 0
    
    for (col, a_list) in zip(cols, arrivals):
        for shukka in df_calend[col]:
            idx = kadoubi_2years.index(shukka) #list内の順番を検索抽出
            arrival_culc = kadoubi_2years[idx + option_day] #着日算出

            if arrival_culc in chakubi_2years:
                a_list.append(arrival_culc)
            else:
                while arrival_culc not in chakubi_2years:
                    i += 1
                    arrival_culc = kadoubi_2years[idx + option_day + i] 
                a_list.append(arrival_culc)
                i = 0
    
    #書式変更　年を外して月日
    arrival_a2 = []
    arrival_b2 = []
    arrival_ex2 = []
    arrival_302 = []

    arraivals_change = [arrival_a, arrival_b, arrival_ex, arrival_30]
    arraivals2 = [arrival_a2, arrival_b2, arrival_ex2, arrival_302]
    
    for (a_list, a_list2) in zip(arraivals_change, arraivals2):
        for chakubi in a_list:
            chakubi = datetime.datetime.strptime (chakubi, '%Y-%m-%d')
            chakubi = chakubi.strftime('%m/%d')
            a_list2.append(chakubi)
    
    #df化
    df_output = pd.DataFrame({
        '受注日': df_calend['受注日'],
        'A (レギュラー)': arrival_a2,
        'B (下記参照)': arrival_b2,
        'SEOTO-EX': arrival_ex2,
        'C（納期30日）': arrival_302,

    })

    #頭の0を無くする
    a_list = []
    b_list = []
    ex_list = []
    c_list = []

    cols2 = ['A (レギュラー)', 'B (下記参照)', 'SEOTO-EX', 'C（納期30日）']
    lists = [a_list, b_list, ex_list, c_list]

    for (col, lst) in zip(cols2, lists):
        for chakubi in df_output[col]:
            split_c = chakubi.split('/')
            month = split_c[0].lstrip("0")
            day = split_c[1].lstrip("0")
            c_without_0 =  month + "/" + day
            lst.append(c_without_0)
    
    df_comp = pd.DataFrame(list(zip(df_output['受注日'], a_list, b_list, ex_list, c_list)), \
                           columns=df_output.columns, index=df_output.index)
    
    return df_comp

def generate_pdf_noncol():

    cols_nonex = ['Aパターン', 'Bパターン', '30日']
    
    #表が格子状になっている場合 lattice=True そうでない　stream=True　複数ページ読み込み pages='all'
    df_calend = df[0]
    #カラムを絞る
    # df_calend = df_calend.drop('KX250AX\rKX260AX', axis=1)
    
    df_calend = df_calend.drop(df_calend.columns[[1, 5, 6, 7]], axis=1) #40日から右カラムの削除
    df_calend = df_calend.dropna(how='any')
    df_calend = df_calend.rename(columns={'Unnamed: 0': '受注日'})
    df_calend = df_calend.loc[1:] #0行目受注日という文字列が入っている


    #曜日を消す
    for col in cols_nonex:
        df_calend[col] = df_calend[col].str[:-2]
    
    #2022年の追加
    for col in cols_nonex:
        df_calend[col] = f'{this_year}年' + df_calend[col]
    
    #datetime型に変換
    for col in cols_nonex:
        df_calend[col] = pd.to_datetime(df_calend[col], format='%Y年%m月%d日')
    
    #年またぎ対応
    #today datetime64型の用意　＊一度dfにしてから変換
    today = datetime.date.today()
    temp_list = []
    temp_list.append(today)
    df_temp = pd.DataFrame(temp_list, columns=['date'])
    df_temp['date2'] = pd.to_datetime(df_temp['date'], format='%Y-%m-%d')
    today2 = df_temp['date2'][0]

    #todayより前のものは年を来年に変換
    for col in cols_nonex:
        df_calend[col] = df_calend[col].map(lambda x: x + relativedelta(years=1) if x < today2 else x)

    #　str型へ　時間を消す
    # Series型にはdtというアクセサが提供されており、Timestamp型を含む日時型の要素から、
    # 日付や時刻のみの要素へ一括変換できる。
    for col in cols_nonex:
        df_calend[col] = df_calend[col].dt.strftime('%Y-%m-%d')
    
    #着日計算
    arrival_a = []
    arrival_b = []
    arrival_30 = []

    arrivals = [arrival_a, arrival_b, arrival_30]

    i = 0
    
    for (col, a_list) in zip(cols_nonex, arrivals):
        for shukka in df_calend[col]:
            idx = kadoubi_2years.index(shukka) #list内の順番を検索抽出
            arrival_culc = kadoubi_2years[idx + option_day] #着日算出

            if arrival_culc in chakubi_2years:
                a_list.append(arrival_culc)
            else:
                while arrival_culc not in chakubi_2years:
                    i += 1
                    arrival_culc = kadoubi_2years[idx + option_day + i] 
                a_list.append(arrival_culc)
                i = 0
    
    #書式変更　年を外して月日
    arrival_a2 = []
    arrival_b2 = []
    arrival_302 = []

    arraivals_change = [arrival_a, arrival_b, arrival_30]
    arraivals2 = [arrival_a2, arrival_b2, arrival_302]
    
    for (a_list, a_list2) in zip(arraivals_change, arraivals2):
        for chakubi in a_list:
            chakubi = datetime.datetime.strptime (chakubi, '%Y-%m-%d')
            chakubi = chakubi.strftime('%m/%d')
            a_list2.append(chakubi)
    
    #df化
    df_output = pd.DataFrame({
        '受注日': df_calend['受注日'],
        'A (レギュラー)': arrival_a2,
        'B (下記参照)': arrival_b2,
        'SEOTO-EX': arrival_b2,
        'C（納期30日）': arrival_302,

    })

    #頭の0を無くする
    a_list = []
    b_list = []
    ex_list = []
    c_list = []

    cols2 = ['A (レギュラー)', 'B (下記参照)', 'SEOTO-EX', 'C（納期30日）']
    lists = [a_list, b_list, ex_list, c_list]

    for (col, lst) in zip(cols2, lists):
        for chakubi in df_output[col]:
            split_c = chakubi.split('/')
            month = split_c[0].lstrip("0")
            day = split_c[1].lstrip("0")
            c_without_0 =  month + "/" + day
            lst.append(c_without_0)
    
    df_comp = pd.DataFrame(list(zip(df_output['受注日'], a_list, b_list, ex_list, c_list)), \
                           columns=df_output.columns, index=df_output.index)
    
    return df_comp



def generate_pdf_nonkxdate():

    cols_nonex = ['Aパターン', 'Bパターン', '30日']
    
    #表が格子状になっている場合 lattice=True そうでない　stream=True　複数ページ読み込み pages='all'
    df_calend = df[0]
    #カラムを絞る
    df_calend = df_calend.drop('KX250AX\rKX260AX', axis=1)
    
    df_calend = df_calend.drop(df_calend.columns[[4, 5, 6]], axis=1) #40日から右カラムの削除
    df_calend = df_calend.dropna(how='any')
    df_calend = df_calend.rename(columns={'Unnamed: 0': '受注日'})
    df_calend = df_calend.loc[1:] #0行目受注日という文字列が入っている

    #曜日を消す
    for col in cols_nonex:
        df_calend[col] = df_calend[col].str[:-2]
    
    #2022年の追加
    for col in cols_nonex:
        df_calend[col] = f'{this_year}年' + df_calend[col]
    
    #datetime型に変換
    for col in cols_nonex:
        df_calend[col] = pd.to_datetime(df_calend[col], format='%Y年%m月%d日')
    
    #年またぎ対応
    #today datetime64型の用意　＊一度dfにしてから変換
    today = datetime.date.today()
    temp_list = []
    temp_list.append(today)
    df_temp = pd.DataFrame(temp_list, columns=['date'])
    df_temp['date2'] = pd.to_datetime(df_temp['date'], format='%Y-%m-%d')
    today2 = df_temp['date2'][0]

    #todayより前のものは年を来年に変換
    for col in cols_nonex:
        df_calend[col] = df_calend[col].map(lambda x: x + relativedelta(years=1) if x < today2 else x)

    #　str型へ　時間を消す
    # Series型にはdtというアクセサが提供されており、Timestamp型を含む日時型の要素から、
    # 日付や時刻のみの要素へ一括変換できる。
    for col in cols_nonex:
        df_calend[col] = df_calend[col].dt.strftime('%Y-%m-%d')
    
    #着日計算
    arrival_a = []
    arrival_b = []
    arrival_30 = []

    arrivals = [arrival_a, arrival_b, arrival_30]

    i = 0
    
    for (col, a_list) in zip(cols_nonex, arrivals):
        for shukka in df_calend[col]:
            idx = kadoubi_2years.index(shukka) #list内の順番を検索抽出
            arrival_culc = kadoubi_2years[idx + option_day] #着日算出

            if arrival_culc in chakubi_2years:
                a_list.append(arrival_culc)
            else:
                while arrival_culc not in chakubi_2years:
                    i += 1
                    arrival_culc = kadoubi_2years[idx + option_day + i] 
                a_list.append(arrival_culc)
                i = 0
    
    #書式変更　年を外して月日
    arrival_a2 = []
    arrival_b2 = []
    arrival_302 = []

    arraivals_change = [arrival_a, arrival_b, arrival_30]
    arraivals2 = [arrival_a2, arrival_b2, arrival_302]
    
    for (a_list, a_list2) in zip(arraivals_change, arraivals2):
        for chakubi in a_list:
            chakubi = datetime.datetime.strptime (chakubi, '%Y-%m-%d')
            chakubi = chakubi.strftime('%m/%d')
            a_list2.append(chakubi)
    
    #df化
    df_output = pd.DataFrame({
        '受注日': df_calend['受注日'],
        'A (レギュラー)': arrival_a2,
        'B (下記参照)': arrival_b2,
        'SEOTO-EX': arrival_b2,
        'C（納期30日）': arrival_302,

    })

    #頭の0を無くする
    a_list = []
    b_list = []
    ex_list = []
    c_list = []

    cols2 = ['A (レギュラー)', 'B (下記参照)', 'SEOTO-EX', 'C（納期30日）']
    lists = [a_list, b_list, ex_list, c_list]

    for (col, lst) in zip(cols2, lists):
        for chakubi in df_output[col]:
            split_c = chakubi.split('/')
            month = split_c[0].lstrip("0")
            day = split_c[1].lstrip("0")
            c_without_0 =  month + "/" + day
            lst.append(c_without_0)
    
    df_comp = pd.DataFrame(list(zip(df_output['受注日'], a_list, b_list, ex_list, c_list)), \
                           columns=df_output.columns, index=df_output.index)
    
    return df_comp


def to_excel(df):
    #メモリー上でバイナリデータを処理       
    output = BytesIO()
    #df化
    df.to_excel(output, index = False, sheet_name='Sheet1')
    #メモリ上から値の取得
    processed_data = output.getvalue()

    return processed_data


if __name__ == '__main__':

    # 内閣府から祝日データを取得、更新したいときにTrueにする
    #指定されたURLからファイルをダウンロードし、ローカルに保存
    if True:
        urllib.request.urlretrieve(public_holiday_csv_url, public_holiday2)

    #############################this year
    # 稼働日　todayから日付を回す
    date_this = datetime.datetime(this_year, 1, 1)
    date_next = datetime.datetime(next_year, 1, 1)
    end_this = datetime.datetime(this_year, 12, 31)

    #今年リスト作成（稼働日/着日）
    while date_this.year ==  this_year :
        #稼働日リスト作成
        if get_kadoubi(date_this):
            kadoubi_this.append(date_this.strftime("%Y-%m-%d"))
        
        # 着日リスト作成
        if get_chakubi(date_this):
            chakubi_this.append(date_this.strftime("%Y-%m-%d"))

        date_this += datetime.timedelta(days=1)
    
    #################################next year
    #来年リスト作成（稼働日/着日）
    while date_next.year == next_year:
        #稼働日リスト作成
        if get_kadoubi(date_next):
            kadoubi_next.append(date_next.strftime("%Y-%m-%d"))
        
        # 着日リスト作成
        if get_chakubi(date_next):
            chakubi_next.append(date_next.strftime("%Y-%m-%d"))

        date_next += datetime.timedelta(days=1)

    kadoubi_2years = kadoubi_this + kadoubi_next
    chakubi_2years = chakubi_this + chakubi_next 

    st.markdown('#### SEOTO-EX 日付あり/なし選択')
    hizuke = st.selectbox('日付あり/なし選択', ['--', '列名なし/Ｂパターン' ,'列名あり/日付なし/Ｂパターン', '列名あり/日付あり'], key='hizuke')

    if hizuke == '--':
        st.info('項目を選択してください')
        st.stop()
    
    elif hizuke== '列名なし/Ｂパターン':
        df_comp = generate_pdf_noncol()
        # to_excel(df_comp)
        df_xlsx = to_excel(df_comp)
        st.download_button(label='Download Excel file', data=df_xlsx, file_name= 'calender.xlsx')

    elif hizuke== '列名あり/日付なし/Ｂパターン':
        df_comp = generate_pdf_nonkxdate()
        # to_excel(df_comp)
        df_xlsx = to_excel(df_comp)
        st.download_button(label='Download Excel file', data=df_xlsx, file_name= 'calender.xlsx')

    elif hizuke == '列名あり/日付あり':
        df_comp = generate_pdf()
        # to_excel(df_comp)
        df_xlsx = to_excel(df_comp)
        st.download_button(label='Download Excel file', data=df_xlsx, file_name= 'calender.xlsx')

    st.markdown('###### 注意！！　GW、お盆、年末年始等が絡む期間 要確認')

    link = '[home](https://cocosan1-hidastreamlit4-linkpage-7tmz81.streamlit.app/)'
    st.markdown(link, unsafe_allow_html=True)
    st.caption('homeに戻る')