#! バッグと財布のサイズに対応
from ast import keyword
from cgitb import text
from concurrent.futures import thread
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import pandas as pd
import random
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import re #正規表現を使うため - テキストの部分一致を取得するため
from tkinter import messagebox #_ to use a messagebox
import sys
import datetime
import math
# from tqdm import tqdm
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from functions import module_01

#- windowsとmacの互換性のあるディレクトリを指定できるようにimport -20220530
from pathlib import Path
from selenium.common.exceptions import TimeoutException

#- 20220521 続いて、小さなウィンドウを表示させない設定をします。
# https://tinyurl.com/y2rrtjh5
import tkinter as tk
root = tk.Tk()

# The content of the messagebox. Return => True or False
is_category_men = messagebox.askyesno(title = 'Confirm', message = '''
Is gender category "Men"??
★★★検索用pprice_range: 要価格帯記入!!★★★
★★★検索用iitem_category: 要商品カテゴリー記入!!★★★
検索用ffile_name:
''',icon='warning')

#- 20220511
if is_category_men == True:
    category_gender = 'Men'
else:
    category_gender = 'Women'

# iitem_category
item_category = 'Bag'

#- 20220521 続いて、小さなウィンドウを表示させない設定をします。
root.withdraw() #! 小さなウィンドウを表示させない

#- windows変更 \を/に変更した- 20220530
#^ current pathを取得 - ここを元にpathを指定していく
#! パスを取得する関数を利用した場合、パスの区切り文字は「\」バックスラッシュになります。
#! .replace(os.sep,'/')で対応する!!
print(os.sep)
cwd = Path.cwd()

#- 相対パスで取得したいファイルのディレクトリのを区切り文字を / に変更し、かつパスオブジェクトから文字列に変換して取得する関数
def neutralize_path_str(path):
    #- パスを繋げ、パスオブジェクトから文字列に変換する
    new_path_str = str(cwd / path)
    #- パス同士をつなげ完成したファイルのディレクトリの区切り文字を / に変更
    return new_path_str.replace(os.sep,'/')

#- 相対パスで取得したいファイルディレクトリをオブジェクトで取得する
def neutralize_path_obj(path):
    #- パスを繋げ、パスオブジェクトを作成
    new_path_obj = cwd / path
    #^ TypeError: replace() takes 2 positional arguments but 3 were given
    #! 不可 オブジェクトにreplace()を使おうとすると上記のエラーが発生する
    # return new_path.replace(os.sep,'/')
    return new_path_obj

user_agent = ['Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_5) AppleWebKit/605.1.15 (KHTML, like Gecko)','Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/605.1.15 (KHTML, like Gecko)','Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15','Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.2 Safari/605.1.15']

t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')
now = datetime.datetime.now(JST)
today = now.strftime('%Y%m%d')
#example 202111041737 分まで
todayToMinute = now.strftime('%Y%m%d%H%M')

#example 2021-11-04 17:37:15 まで
now = now.strftime('%Y-%m-%d %H:%M:%S')
print('\n', 'スタート時刻: ',now, '\n')
# print(repr(now))

options = webdriver.ChromeOptions()
options.add_argument('--user-agent=' + user_agent[random.randrange(0, len(user_agent), 1)])
options.add_argument('--incognito')
options.add_experimental_option('detach', True)

options.add_experimental_option('excludeSwitches', ['enable-logging'])

#! Change executable path according to your chrome webdriver location
chromeDriverLocation = str(list(neutralize_path_obj('tools').glob('*'))[0]).replace(os.sep,'/')

service = Service(chromeDriverLocation)
# service = Service("C:\\chromedriver")
driver = webdriver.Chrome(service=service,options=options)
driver.maximize_window()

#https://tanuhack.com/stable-selenium/#i-2
#implicitly_waitメソッドを設置する場所は、Chrome Driverを起動させた直後で良いでしょう。_20220319
driver.implicitly_wait(10)

#- 20220403 https://office54.net/python/scraping/selenium-wait-time
wait = WebDriverWait(driver, 20)

d_list = []
#! ebayのFXで付加する項目を認識するためのflag
ebay_category =''
iterated_num = 0 #タイトルを一意のものにするために付加する番号
current_item_count = 0 #現在取得中のアイテムの番数を取得 - 20220512
#! for storing pictures url for inserting downloaded photos directly to ebay
DL_pic_list = []

# file_name = '2nd-kateSpade-30k-wallet-men_ver2_'
file_name = ''

#- 途中でスクレイピングが終了してしまった場合に使用するファイル名
discontinued_file_name = file_name

#- 税抜価格
item_price = 0


# #- urlが下記の形式でない場合、webページ上で次へボタンを押してurlを切り替える必要がある
# #- ralph lauren bag men 00k~All
driver.get('https://www.2ndstreet.jp/search?category=951002&keyword=Ralph%20Lauren%20%E3%83%90%E3%83%83%E3%82%B0&conditions[]=N&conditions[]=S&conditions[]=A&conditions[]=B&conditions[]=C')

base_url = 'https://www.2ndstreet.jp/search?category=951002&keyword=Ralph%20Lauren%20%E3%83%90%E3%83%83%E3%82%B0&conditions[]=N&conditions[]=S&conditions[]=A&conditions[]=B&conditions[]=C&page={}'

#- ★ファイル名に記述する価格を記述する
#!           \d+[kK]....    この正規表現に合致するようにする
#- 2nd_sel_ebay_sneakerModel.py にて使用 pprice_range
# price_range='14k~16k'
price_range='00k~All'

iterated_num = 0 #タイトルを一意のものにするために付加する番号

#¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥¥
#-20220307 replace(',','')を加えた <= 1000以上の検索結果のときに
#- 1,000のようにカンマが入り、int()時にエラーが起きるため
#! replace()は、入替対象のものがなかった場合でもエラーは起きずそのままcontinueする
search_result_number = int(driver.find_element_by_css_selector('div#ecResultNum > p > span').text.replace(',',''))
print('='*30)
print('search_result_number:',search_result_number)
print('='*30)

# get the total page number to decide how many loops you need
total_page_number = math.ceil(search_result_number / 60)
print('='*30)
print('total_page_number:',total_page_number)
print('='*30)

#For storing products links
product_link_list= []

#For storing all the products images link
img_links_list=[]

#- 全体の何番目かを確認する用
item_count = 0

#! rang(1,x)と設定しているため、iは1から始まる
for i in range(1,int(total_page_number)+1):

    #- This line will open chrome with provided page number url
    # driver.get('https://www.2ndstreet.jp/search?category=950001&keyword=%s&sortBy=arrival&page=%i'% (searchKeyword,i))
    # driver.implicitly_wait(3)

    print('rangeのi', i, 'もしdiscontinued_file_nameでファイルが作成されるようであれば、sleepの間隔が短すぎる可能性があるため要変更')

    if i > 1:
        #! NG どういうわけか base_urlがformat(i)で更新されない!!
        #! => 一度ページ数1で固定されたbase_urlを使うことになるから.format(i)で
        #! そもそも更新されない
        #! 別の変数を用意する必要あり!!
        # base_url = base_url.format(i)
        url = base_url.format(i)
        driver.get(url)
        sleep(2)
    else:
        pass

    #This will get the whole div in which the products are present
    products_list = driver.find_element(By.CLASS_NAME, 'itemList')

    #This will find all the link tag in the div
    products_links = products_list.find_elements(By.TAG_NAME, 'a')

    #Variable to hold Products links on a specific page
    products_page=[]

    #Looping through all the products to get the images
    for product_link in products_links:

        #After visiting each product the website asks for login so in order to bypass that
        if product_link.get_attribute('href') == 'https://www.2ndstreet.jp/user/login':
            pass
        else:
            products_page.append(product_link.get_attribute('href'))

    #- product is each item page
    for index, product in enumerate(products_page):
        #- ここから - 20220417
        try:
            driver.get(product)
            # sleep(2)

            print('\n')
            print('='*20,'2: to check where the code stops.','='*20)

            #-強制的にエラーを起こしている - 20220531
            try:
                #element_present = WebDriverWait(driver, 0.1).until(EC.presence_of_element_located((By.ID, "viewHistory")))
                driver.set_page_load_timeout(20)
                print('WebDriverWait ok')
            except TimeoutException:
                print("Timed out waiting for page to load")
                raise Exception("Exception from try line 226")

            #- 秒数を1秒に変更 - アイテムの取得に平均6秒かかっていたため
            sleep(1)
            item_index = index #! indexに影響を与えないようprint用にindexのコピーを作成

            if (index + 1) % 60 == 0:
                item_count += 60
                #! item_count + (item_index+1) が60の倍数の際に、60余計に加算されて計算されてしまう
                #! 上記の仕様でアイテムカウントがおかしくなるのを防ぐために作成
                item_index = -1

            # print('='*10,'アイテムカウント: ',item_count + (item_index+1),' / 送料抜き価格: ',item_price,'='*10)
            print('='*15,f'{current_item_count + 1}/{search_result_number}番目 | 値幅: {price_range} | category: {category_gender}','='*15)

            item_url = product
            iterated_num += 3
            current_item_count += 1
            page_soup = BeautifulSoup(driver.page_source, 'lxml')
            zoom_pictures = []
            pictures = page_soup.select('div[class="goods_popname"] > ul > li')

            for i, picture in enumerate(pictures):
                picture = picture.select_one('img').get('data-src')
                zoom_pictures.append(picture)
                #¥ storing for downloaded pictures
                DL_pic_list.append(picture)
            #- 新しいリストを作成 <= img_links
            zoom_pictures = '|'.join(zoom_pictures)
            # DL_pic_list = ','.join(DL_pic_list)

            #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            #- 修正 - 2ndStreetのwebsiteが p => divに変更した? - 20220406
            #- .productNameは一つしかないので、要素は指定しないこととする?
            # item_title = page_soup.select_one('p.productName').text
            item_title = page_soup.select_one('div.productName').text

            # category  不要

            #- 関数1
            def purify_size(x,y):
                if page_soup.find(text=re.compile(x)):
                    try:
                        size_cm = float(page_soup.find(text=re.compile(x)).replace(x,'').replace('/',''))
                        return size_cm
                    #example: 幅:24-36 <= string so float() cannot be used
                    #caused ValueError
                    except ValueError as e:
                        #! string形式で代入し、inchは''で表示させる
                        #厚み / <= / を取り除くための2つ目のreplace()
                        size_cm = page_soup.find(text=re.compile(x)).replace(x,'').replace('/','')
                        print('exceptのエラーメッセージ: ',e)
                        print('except ValueError(String型): ',size_cm)
                        return size_cm
                elif page_soup.find(text=re.compile(y)):
                    try:
                        size_cm = float(page_soup.find(text=re.compile(y)).replace(y,'').replace('/',''))
                        return size_cm
                    except ValueError as e:
                        #! string形式で代入し、inchは''で表示させる
                        size_cm = page_soup.find(text=re.compile(y)).replace(y,'').replace('/','')
                        print('exceptのエラーメッセージ: ',e)
                        print('except ValueError(String型): ',size_cm)
                        return size_cm
                else:
                    size_cm = ''
                    return size_cm

            #- 関数2
            def to_actual_size(handle, depth, height, width, unit):
                result = 'Approx size(handle / depth / height / width: ' + str(handle) + unit + ' / ' + str(depth) + unit + ' / ' + str(height) + unit + ' / ' + str(width) + unit + ')'
                return result

            handle_cm = purify_size('持ち手 ','no-css_selector')
            # size_handle_cm = float(page_soup.find(text=re.compile('持ち手')).replace('持ち手 ','' ))

            depth_cm = purify_size('マチ ','厚み ')

            height_cm = purify_size('高さ ','縦 ')

            width_cm = purify_size('幅 ','横 ')

            handle_inch = round(handle_cm / 2.54, 2) if handle_cm and type(handle_cm)== float  else ''
            depth_inch = round(depth_cm / 2.54, 2) if depth_cm and type(depth_cm)== float  else ''
            height_inch = round(height_cm / 2.54, 2) if height_cm and type(height_cm)== float  else ''
            width_inch = round(width_cm / 2.54, 2) if width_cm and type(width_cm)== float else ''

            actual_size_cm = to_actual_size(handle_cm, depth_cm, height_cm, width_cm, 'cm')

            actual_size_inch = to_actual_size(handle_inch, depth_inch, height_inch, width_inch, 'inch')

            #- 修正 - 2ndStreetのwebsiteが p => divに変更した? - 20220406
            #- .productNameは一つしかないので、要素は指定しないこととする?
            # item_title = page_soup.select_one('p.productName').text
            bland_name = page_soup.select_one('div.blandName').text.strip()
            # product_name = page_soup.select_one('p.productName').text.replace('/',' ').strip()
            product_name = page_soup.select_one('div.productName').text.replace('/',' ').strip()
            item_title = bland_name+' '+product_name

            # category  不要
            if page_soup.select_one('dt:-soup-contains("型番")+dd'):
                model = page_soup.select_one('dt:-soup-contains("型番")+dd').text.strip()
                model='' if model == 'ー' else model

                color = page_soup.select_one('dt:-soup-contains("カラー")+dd').text.strip()

                pattern = page_soup.select_one('dt:-soup-contains("柄")+dd').text.strip()

                material = page_soup.select_one('dt:-soup-contains("素材・生地")+dd').text.strip()
                material='' if material == 'ー' else material

                category_complete = []
                category_whole = page_soup.select('section#breadcrumb > span > a')
                category_brand = page_soup.select_one('section#breadcrumb > span:nth-of-type(3) > a > span').text
                for category in category_whole:
                    category_complete += [category.text]

                category_complete = ' '.join(category_complete)
                category_complete = category_complete.replace('ホーム','').replace(category_brand,'').replace('商品詳細','').replace('バッグ','',1).strip()

                brand_name = category_brand

                #! decompose 指定した要素を完全に削除し破壊するメソッド いらない子要素を削除
                # https://tinyurl.com/ycd9sqr3
                page_soup.select_one('div.conditionStatus > div').decompose()

                #- 目的のもののみ取得できた! 20220220
                condition = page_soup.select_one('div.conditionStatus').text.strip().replace( '\n' , '').replace('商品の状態: ','')

                # print('decompose後のcondition', condition)

                try:
                    #- split('店頭')で余分な言葉を省いている
                    condition_detail = page_soup.select_one('div#shopComment').text.strip().split('店頭')[0]
                except AttributeError as e:
                    print('ErrorMessage:',e)
                    condition_detail = ''

                #- 属性の値を取得
                item_price = int(page_soup.select_one('span.priceNum').get('content').strip())

                d_list.append({
                    '*Action(SiteID=US|Country=JP|Currency=USD|Version=745|CC=UTF-8)': 'Add',
                    'DL-eb-category': 'bag_m:52357 / bag_w:169291',
                    'X-number': "'"+str(iterated_num).zfill(3),
                    '*Title': '',
                    # 'U-title': item_title,
                    'U-title': item_title,
                    'X-category':category_complete,
                    'Ref-title_category_combined':'Dont use this time',
                    'Ref-translation':'^^=if(E2="","",googletranslate(E2,"ja","en"))',
                    #- モデル名の参照を削除-すでにタイトルに含まれているため
                    'URef-title-completed':'^^=H2&" "&"Bag"&" "&"from Japan"&" "&C2',
                    'Ref-len()': '^^=len(D2)',
                    'Ref-material_translation': '^^=if(Y2="","",googletranslate(Y2,"ja","en"))',
                    # 'URef-size_for_ConditionDescription':'^^=M2&" | "&N2&" | "&" Model: "&X2&" | "&"Material: "&K2&" | "&"Shipping from Japan. For more details, please check the pictures carefully and judge the condition.',

                    'URef-size_for_ConditionDescription':'^^=M2&" | "&N2&" | "&" Model: "&X2&" | "&"Material: "&K2&" | "&"Shipping from Japan. Note: All you receive is what you see in the pictures(a hanger or a stand is not included.). If some parts or accessories are not shown in the pictures, such as a shoulder strap, they are not included. For more details, please check the pictures carefully and judge the condition.',
                    'X-actual_size_inch': actual_size_inch,
                    'X-actual_size_cm': actual_size_cm,
                    'CustomLabel' : 'Please fill in like this. py-Supplier-BRAND-Category-Gender-DATE-Number',
                    '*Description' : '',
                    #¥ New 20220211
                    #- ConditionDescription にはサイズや型番、状態が入る
                    'ConditionDescription': '',
                    '*StartPrice' : '',
                    'MinimumBestOfferPrice': '"¥=R2',
                    'url': item_url,
                    'costPrice + s/p': item_price + 770,
                    # 'combined':'=c{num}+h{num}'.format(num = i+2),
                    'ConditionID' : '3000',
                    'PicURL':zoom_pictures,
                    'X-model': model,
                    'X-material': material,
                    'X-condition': condition,
                    'X-condition_detail': condition_detail,
                    'C:Brand': brand_name,
                    # 'C:Brand': delete_brackets.delete_brackets(brand_name),
                    'X-color': color,
                    'X-gender': '',
                    # 'X-all_sizes': store_all_sizes,
                })
                #- エラー等でコードが中断されたように途中経過をexcelファイルに書きだし記録: 20220528
                #- 指定したアイテム取得毎にバックアップとしてファイルを作成
                num_for_backup = 5
                if current_item_count % num_for_backup == 0:
                # trf-jilSander-30k-men-20220428
                    file_name = f'2nd-{brand_name}-{current_item_count}番目まで{price_range}{item_category}-{category_gender}-'

                    df = pd.DataFrame(d_list)

                    if is_category_men == True:
                        df_add = pd.read_excel('e_category_no/athleticShoes_men.xlsx')
                    else:
                        df_add = pd.read_excel('e_category_no/athleticShoes_women.xlsx')

                    df_complete = pd.concat([df,df_add],axis=1)
                    # df_complete = df_complete

                    #- １つ目の引数=行、2つ目の引数=列を指定して代入
                    # https://tinyurl.com/2e7zg56s
                    df_complete.loc[0,'*C:US Shoe Size'] = '^^=M2'

                    #- backupの場所を指定する
                    backup_location = neutralize_path_str('csv_folder/csv_for_concat/ebay/z_backup')
                    print('\n', 'backup_location: ',backup_location)

                    #! change the file name every time
                    #-20220530
                    df_complete.to_excel(f'{backup_location}/{file_name}{todayToMinute}.xlsx', index=False)

                    #^ バックアップファイルが１つ目のときは処理をスキップ-消すファイルがないため
                    if current_item_count - num_for_backup != 0:
                        #- 前回作成したファイルが何番目だったかを特定する
                        previous_item_count = current_item_count - num_for_backup

                        previous_file_name =  f'2nd-{brand_name}-{previous_item_count}番目まで{price_range}{item_category}-{category_gender}-'

                        #- 古いファイルを削除する※rmdir()とunlink()の使い分け
                        Path(f'{backup_location}/{previous_file_name}{todayToMinute}.xlsx').unlink()

                        print('\n', '★削除したファイル :', f'{backup_location}/{previous_file_name}{todayToMinute}.xlsx','\n')

                        print('\n', '★途中経過記録保存場所 - concat_refへ :', f'{backup_location}{file_name}{todayToMinute}.xlsx','\n')

            else:
                #! elseのif どのくらいの頻度で倉庫取り寄せ品があるか知るために
                #! あえてexcelファイルに出力するようにしている。
                break
                d_list.append({
                    '*Action(SiteID=US|Country=JP|Currency=USD|Version=745|CC=UTF-8)': 'Dont Use this item. Delete this row',
                    'DL-eb-category': 'Dont Use this item. Delete this row',
                    'X-number': str(iterated_num).zfill(3),
                    '*Title': 'Dont Use this item. Delete this row.Dont Use this item. Delete this row.Dont Use this item. Delete this row',
                    # 'U-title': item_title,
                    'U-title': 'Dont Use this item. Delete this row',
                    'X-category':'Dont Use this item. Delete this row',
                    'Ref-title_category_combined':'Dont Use this item. Delete this row',
                    'Ref-translation':'Dont Use this item. Delete this row',
                    #- モデル名の参照を削除-すでにタイトルに含まれているため
                    'URef-title-completed':'Dont Use this item. Delete this row',
                    'Ref-material_translation': 'Dont Use this item. Delete this row',
                    'URef-size_for_ConditionDescription':'Dont Use this item. Delete this row',
                    'X-actual_size_inch': 'Dont Use this item. Delete this row',
                    'X-actual_size_cm': 'Dont Use this item. Delete this row',
                    'CustomLabel' : 'Dont Use this item. Delete this row',
                    '*Description' : 'Dont Use this item. Delete this row',
                    #¥ New 20220211
                    #- ConditionDescription にはサイズや型番、状態が入る
                    'ConditionDescription': 'Dont Use this item. Delete this row',
                    '*StartPrice' : 'Dont Use this item. Delete this row',
                    'MinimumBestOfferPrice': '"¥=R2',
                    'url': item_url,
                    'costPrice + s/p': 0,
                    # 'combined':'=c{num}+h{num}'.format(num = i+2),
                    'ConditionID' : 'Dont Use this item. Delete this row',
                    'PicURL':'Dont Use this item. Delete this row',
                    'X-model':'Dont Use this item. Delete this row',
                    'X-material': 'Dont Use this item. Delete this row',
                    'X-condition': 'Dont Use this item. Delete this row',
                    'X-condition_detail': 'Dont Use this item. Delete this row',
                    'C:Brand': 'Dont Use this item. Delete this row',
                    # 'C:Brand': delete_brackets.delete_brackets(brand_name),
                    'X-color': 'Dont Use this item. Delete this row',
                    'X-gender': 'Dont Use this item. Delete this row',
                    # 'X-all_sizes': store_all_sizes,
                    #¥ 追加_20220221
                    'DL-pic': '',
                })
                print('='*30)
                print('除外: 倉庫より出荷: ',d_list[-1])

        except Exception as exception:
            print('\n', '読み込みエラー? Exception: ',exception,'\n')

#- 検索要ffile_name
# file_name = f'2nd-{brand_name}-sneaker-{category_gender}-'
file_name = f'2nd-{brand_name}-{price_range}{item_category}-{category_gender}-'

#It will create a new file inside the project
df = pd.DataFrame(d_list)

if is_category_men == True:
    df_add = pd.read_excel('e_category_no/bag_men.xlsx')
else:
    df_add = pd.read_excel('e_category_no/bag_women.xlsx')

df_complete = pd.concat([df,df_add],axis=1)

#- 20220530 windows、Macに対応できるpathに変更
save_location = f'{str(cwd)}/csv_folder/csv_for_concat/ebay/'

df_complete.to_excel(f'{save_location}{file_name}{todayToMinute}.xlsx', index=False)

print('\n', 'エラーなし保存場所 - concat_refへ :', f'{save_location}{file_name}{todayToMinute}.xlsx','\n')

driver.close()

# folder_dir = neutralize_path_str('results/ebay')
folder_dir = save_location
print('folder_dir: ', folder_dir)

#- プログラム終了後、保存先フォルダーを開く  https://tinyurl.com/25ojjpdl
module_01.open_folder(folder_dir)






