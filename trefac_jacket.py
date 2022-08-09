from tabnanny import filename_only
from time import sleep
from numpy import delete
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import requests #! いらない
from bs4 import BeautifulSoup
import pandas as pd
import random
import datetime
import glob
# 括弧内及び括弧内の文字を削除する関数をimport
#^ パスの指定方法: https://tinyurl.com/ybfxc7av
# from functions import delete_brackets
#- functions > delete_bracketsファイルの中から関数delete_bracketsをimportしている
from functions.delete_brackets import delete_brackets
# _to use a messagebox
from tkinter import messagebox
import sys
import math

import os
#! NEW 20220403 https://office54.net/python/scraping/selenium-wait-time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from functions import module_01
from pathlib import Path
import re
import tkinter as tk

#- 検索word mmesagebox
# The content of the messagebox. Return => True or False
#! ------ここからコメント
# is_category_men = messagebox.askyesno(title = 'Confirm', message = '''Is gender category "Men"??
# 商品の絞り込みで商品の状態で悪いものは除いて検索したurlを使用すること''',icon='warning')

# root = tk.Tk()
# # - 20220521 続いて、小さなウィンドウを表示させない設定をします。
# root.withdraw()  # ! 小さなウィンドウを表示させない

# if is_category_men == True:
#     category_gender = 'Men'
# else:
#     category_gender = 'Women'

is_category_men = True #要変更 popup時は要コメント
category_gender = 'Men' #要変更 popup時は要コメント
#! ------ここまでコメント


t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')
now = datetime.datetime.now(JST)
today = now.strftime('%Y%m%d')
now = now.strftime('%Y-%m-%d %H:%M:%S') # 2022-05-30 17:08:53
print(f'スタート時刻: {now}')

user_agent = ['Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36','Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Safari/537.36','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36']

options = webdriver.ChromeOptions()
# options.add_argument('--headless')
options.add_argument('--user-agent=' + user_agent[random.randrange(0, len(user_agent), 1)])

# options.add_argument('--incognito')
#! ChromeDriver起動時にdetachオプションを設定し、quit()やclose()を呼ばなければ、Selenium実行後もChromeを開いたままにできる。
options.add_experimental_option('detach', True)

#- インスタンス化 - 20220621
control_path_class = module_01.Control_path()

chromeDriverLocation = str(list(control_path_class.neutralize_path_obj('tools').glob('*'))[0]).replace(os.sep,'/')

service = Service(chromeDriverLocation)

driver = webdriver.Chrome(service=service, options=options)
# driver.maximize_window()

# #- ウェブサイトが呼び込むのを何秒待つかを指定できる
driver.implicitly_wait(10)

#検索用word uurl 20220713
#test用
fixed_url_part = ' https://www.trefac.jp/store/t1c92psb/?srchword=North%20Face&selects=5000&selecte=14000&gcondition=6,5,4,3&step=1&disp_num=90 '


#- uurl検索要url
# fixed_url_part = ' https://www.trefac.jp/store/t1c92psb/?srchword=Ferrari&gcondition=6,5,4,3&step=1&disp_num=90 '

fixed_url_part = fixed_url_part.strip() #-前後の空白を削除する-20220803

driver.get(f'{fixed_url_part}')
sleep(2)

# 1件のみ
# driver.get('https://www.trefac.jp/store/search_result.html?srchword=&step=1&k_uid=CpaqJfdnXxO0ukWZWyJfup0Lmsbq7I&q=%E3%82%B7%E3%83%A7%E3%83%BC%E3%83%AB%E3%82%AB%E3%83%A9%E3%83%BC%E3%83%8D%E3%83%83%E3%83%97%E5%85%A5%E3%82%8A%E3%82%B8%E3%83%A3%E3%82%B1%E3%83%83%E3%83%88&searchbox=1')
# sleep(2)

d_list = []

iterated_num = 0 #タイトルを一意のものにするために付加する番号

#! 検索するアイテムごとに修正が必要
#! 34行目のコードに &key={} をurlの最後に付け足す必要あり!
#to 45行目のrange()の数を取得するページ数に応じて変更の必要あり!
# base_url ='https://www.trefac.jp/store/t1c92psb/?srchword=Drake%27s&gcondition=6,5,4,3&step=1&order=pdown&key={}'

base_url =fixed_url_part + '&key={}'

# 'https://www.yourshoppingmap.com/brand/295-moorer?page=2#boutiques'

#- 商品の状態のランク - 20220621
condition_ranks = {
    'S': 'Condition：Unused.',
    'A': 'Condition：Near Mint - Pre-owned, almost no signs of use.',
    'B': 'Condition：Between Excellent ~ Very Good - Pre-owned, no noticeable damages ~  Very Good - with some signs of use, possibly some stains or some slight damages.',
    'C': 'Condition：Very good - Very good condition, with some signs of use, possibly some stains or some slight damages.'
}

#- 関数_1
def strip_text(css_selector):
    result = page_soup.select_one(css_selector).text.strip()
    return result

#- 関数_2
def purify_size(x,y=''):
    if page_soup.select_one(x):
        size_cm = page_soup.select_one(x)
        # print('\n', 'size_cm:',size_cm  ,'\n')
        try:
            #! オンラインショップ上のサイズ表記が間違っており "約c33m" と記載されており、エラーになったため修正
            size_cm = float(size_cm.text.replace('約',' ').replace('c', '').replace('m', '')) if size_cm else ''
            #! returnをつけ忘れないように! そうしないと計算した値を後から代入できない
            return size_cm

        except ValueError as e:
            print('catch ValueError:', e)
            #! 直でreturn 可能では? return float(size_cm.text.replace(...))
            #¥ 下記のように空白が必要か結果を確認する
            size_cm = ''
            return size_cm
    elif page_soup.select_one(y):
        size_cm = page_soup.select_one(y)
        # print('\n', 'size_cm:',size_cm  ,'\n')
        try:
            size_cm = float(size_cm.text.replace('約',' ').replace('c', '').replace('m', '')) if size_cm else ''
            return size_cm
        except ValueError as e:
            print('catch ValueError:', e)
        size_cm = ''
        return size_cm
    else:
        size_cm = ''
        # print('\n', 'size_cm:',size_cm  ,'\n')
        return size_cm

#- 関数_3
# def to_actual_size(body_width_cm, shoulder_cm,length_cm,unit):
#     result = 'Approx size(body_width / shoulder / length: ' + str(body_width_cm) + unit + ' / ' + str(shoulder_cm) + unit + ' / ' + str(length_cm) + unit + ')'
#     return result

# def to_actual_size(body_width_cm, shoulder_cm,length_cm,unit):
def to_actual_size(body_width_cm, shoulder_cm, sleeve_length_cm,length_cm,unit):
    result = f'Approx size(body width: {str(body_width_cm)}{unit} / shoulder_width: {str(shoulder_cm)}{unit}  / sleeve_length: {str(sleeve_length_cm)}{unit}  /length: {str(length_cm)}{unit} )'
    return result

#- 関数4 - 20220621 複数行の文字列からモデル等特定の文字列を取得する関数
def regex_find_all_to_one(pattern, item_detail):
    try:
        result = re.findall(pattern,item_detail)[0]
        return result
    except Exception as e:
        print(f'★例外-regex_find_all_to_one(pattern, item_detail): {e}')
        return 'NA'

# get total number of the items in all pages
search_result_number = int(driver.find_element_by_css_selector('span.search_result_num').text)
print('='*30)
print('search_result_number:',search_result_number)
print('='*30)

# get total number of the items in the current page
current_item_number = int(driver.find_element_by_css_selector('span#displayoption_num_btn_current').text.replace('件',''))
print('='*30)
print('current_item_number:',current_item_number)
print('='*30)

# get the total page number to decide how many loops you need
total_page_number = math.ceil(search_result_number / current_item_number)
print('='*30)
print('total_page_number:',total_page_number)
print('='*30)

for i in range(total_page_number):
    # print('='*30, i, '='*30)
    url = base_url.format(i+1)
    #- seleniumで読み込み

    if i == 0:
        pass
    else:
        driver.get(url)
        sleep(4)

    soup = BeautifulSoup(driver.page_source, 'lxml')
    a_tags = soup.select('li.p-itemlist_item > a')

    for i, a_tag in enumerate(a_tags):
        print('='*10,i,'='*10,)
        #¥ 一意のタイトルにするために付属させる番号

        iterated_num += 3
        # if i < 2:
        page_url = a_tag.get('href')
        #- 毎回ここでa_tagが切り替わるたびにそのaタグのurlを読み込む
        driver.get(page_url)
        sleep(2)
        #¥ seleniumの処理 zoom画像の抜き出し
        thumbnails = driver.find_elements_by_css_selector('li.gdimage_thumb_list_item')

        #- zoomした写真をひとまとまりにする必要あり url|url|url|
        zoom_pictures = []

        #! 最初の一枚はもとから拡大写真のため写真をクリックする操作が不要
        for i, thumbnail in enumerate(thumbnails):
            if i==0:
                zoom_picture = driver.find_element_by_css_selector('div.zoomPad > img').get_attribute('src')
                sleep(0.5)
                zoom_pictures.append(zoom_picture)
            else:
                thumbnail.click()
                sleep(0.5)
                zoom_picture = driver.find_element_by_css_selector('div.zoomPad > img').get_attribute('src')
                zoom_pictures.append(zoom_picture)

        #- わかりやすいように変数名を soup から page_soupに変えた
        page_soup = BeautifulSoup(driver.page_source, 'lxml')

        #! 複数のセレクタで対象を絞り込む(区切り文字なし) 対象の2つのクラスを持つもののみ(間にスペースを入れない)を指定 strip()で空白削除 '.gdbrand.p-typo_body1_a'
        #^ 関数_1使用
        brand_name = strip_text('.gdbrand.p-typo_body1_a > a')
        print('ブランド名: ', brand_name)
        item_name = strip_text('.gdname.p-typo_head3_a')

        #- タグ表示サイズ取得
        tag_size = page_soup.select_one('p.gdsize.p-typo_body1_d')
        # print('\n', 'tag_size01: ', tag_size ,'\n')
        # tag_size = tag_size.text if tag_size else ''
        tag_size = tag_size.text.replace('タグ表記サイズ：','').strip() if tag_size else '' # S（ 参考サイズ：S ）
        # print('\n', 'tag_size02: ', tag_size ,'\n')
        tag_size = delete_brackets(tag_size) # S
        # print('\n', 'tag_size: ', tag_size ,'\n')

        #- 20220708 update 表記なしという日本語を取得してしまった時用
        pattern = r'[ぁ-んァ-ヶ亜-熙纊-黑]+'

        if re.findall(pattern,tag_size):
            tag_size = 'unknown'

        # brand_name = page_soup.select_one('.gdbrand.p-typo_body1_a > a').text.strip()
        # item_name = page_soup.select_one('.gdname.p-typo_head3_a').text.strip()
        item_title = brand_name + ' ' + item_name
        item_url = page_url
        #! replace()で余分なコンマや文字を消し、最後に整数にした
        item_price = int(page_soup.select_one('.gdprice_main').text.replace('￥','').replace(',','').replace('税込',''))
        shipping_cost = page_soup.select_one('td:-soup-contains(全国一律)').text.strip()

        # url|url|url|url
        zoom_pictures = '|'.join(zoom_pictures)

        # gender = page_soup.select_one('tbody.gddescription_attr_body > tr:first-of-type > td').text.strip()
        gender = strip_text('tbody.gddescription_attr_body > tr:first-of-type > td')

        category_big = strip_text('th:-soup-contains(カテゴリー) + td > a:first-of-type')
        #! new
        # category_big = page_soup.select_one('th:-soup-contains(カテゴリー) + td > a:first-of-type').text.strip()
        # print('big',category_big)

        category_small = page_soup.select_one('th:-soup-contains(カテゴリー) + td > a:nth-of-type(2)').text.replace('/', ' ').strip()

        #- カテゴリーがaタグで２つに別れていたのを一つにまとめた
        category_whole = category_big + ' ' + category_small

        #- update - 20220621
        condition = strip_text('tbody.gddescription_attr_body > tr:nth-of-type(4) > td')

        if '未使用品' in condition:
            condition_rank = condition_ranks['S']
        elif '未使用に近い' in condition:
            condition_rank = condition_ranks['A']
        elif '使用感をあまり感じない' in condition:
            condition_rank = condition_ranks['B']
        elif 'やや傷や汚れがあり' in condition:
            condition_rank = condition_ranks['C']
        else:
            condition_rank = '--'

        print('\n', 'condition_rank: ',condition_rank, '\n')

        # condition = page_soup.select_one('tbody.gddescription_attr_body > tr:nth-of-type(4) > td').text.strip()

        accessories = strip_text('tbody.gddescription_attr_body > tr:nth-of-type(5) > td')

        #- 文字列で要素を取得! - soup.select('li:-soup-contains("Python")')
        #! 事前に変数の用意が必要
        staff_comment = ''
        color = ''
        model = ''
        material = ''
        manufactured = ''
        remark = ''
        condition_detail = ''

        item_detail = page_soup.select_one('p.gddescription_free')
        #- アイテムの詳細がpタグ一つで全て文章で書いてある場合と、tableタグで作られている場合の2パターンがあるので分岐
        #^ - update 情報が表になっていないパターンでもモデル名等を取得できる用にした - 20220621
        if item_detail:
            item_detail = page_soup.select_one('p.gddescription_free').text

            color = regex_find_all_to_one('(?<=【カラー】).+', item_detail)

            model = regex_find_all_to_one('(?<=【型番】).+', item_detail)

            manufactured = regex_find_all_to_one('(?<=【製造国】).+', item_detail)

            material = regex_find_all_to_one('(?<=【素材】).+', item_detail)

            #- 改行があると改行する前までしか情報が取得できず不十分であるため、改行なしの一続きの文章にする - 20220621
            item_detail_no_line = item_detail.replace('\r\n','').replace('\n','')

            #- 不要な文章が入っていることがあるので削除
            remark = regex_find_all_to_one('(?<=【備考】).+', item_detail_no_line)
            remark = remark.replace('こちらの商品は中古品の為以下の状態となっており、使用に伴います細かなキズ、汚れがございます。','').replace('店頭にて展示販売を行っている為、記載に無い細かなキズ、汚れが見受けられる場合がございます。','')

            condition_detail = regex_find_all_to_one('(?<=【状態】).+', item_detail_no_line)
            condition_detail = condition_detail.replace('こちらの商品は中古品の為以下の状態となっており、使用に伴います細かなキズ、汚れがございます。','').replace('店頭にて展示販売を行っている為、記載に無い細かなキズ、汚れが見受けられる場合がございます。','')
            # print('\n', 'remark: ', remark ,'\n')
            # print('\n', 'condition_detail: ', condition_detail ,'\n')

        else:
            staff_comment = page_soup.select_one('th:-soup-contains("スタッフコメント") + td')
            staff_comment = staff_comment.text if staff_comment else ''

            color = page_soup.select_one('th:-soup-contains("カラー")+ td')
            color = color.text if color else ''

            model = page_soup.select_one('th:-soup-contains("型番")+ td')
            model = model.text if model else ''

            material = page_soup.select_one('th:-soup-contains("素材")+ td')
            material = material.text if material else ''

            manufactured = page_soup.select_one('th:-soup-contains("製造国")+ td')
            manufactured = manufactured.text if manufactured else ''

            remark = page_soup.select_one('th:-soup-contains("備考")+ td')
            remark = remark.text if remark else ''

            condition_detail = page_soup.select_one('th:-soup-contains("状態")+ td')
            condition_detail = condition_detail.text if condition_detail else ''

        #-サイズ取得
        body_width_cm = ''
        shoulder_cm = ''
        length_cm = ''
        # handle_cm = ''
        # shoulder_strap_cm = ''

        #^ レアサイズ url - https://tinyurl.com/2brdn2c6
        body_width_cm = purify_size('dt:-soup-contains("身幅") + dd','table.tbl_autosize > tbody > tr:nth-of-type(2) > td.autosize_value')
        # print('body_width_cm: ', body_width_cm)

        shoulder_cm = purify_size('dt:-soup-contains("肩幅") + dd','table.tbl_autosize > tbody > tr:first-of-type > td.autosize_value')
        # print('shoulder_cm: ', shoulder_cm)

        #¥  aaaa
        # length_cm = purify_size('dt:-soup-contains("着丈") + dd','table.tbl_autosize > tbody > tr:nth-of-type(4) > td.autosize_value')

        #- 20220720 update
        sleeve_length_cm = purify_size('dt:-soup-contains("袖丈") + dd','table.tbl_autosize > tbody > tr:nth-of-type(3) > td.autosize_value')

        length_cm = purify_size('dt:-soup-contains("着丈") + dd','table.tbl_autosize > tbody > tr:last-of-type > td.autosize_value')
        # print('length_cm: ', length_cm)

        #! ショルダーバッグの紐の長さが "約87-" となっていることから float()できなかった
        #- in演算子と or演算子を使って条件を記載する
        # if page_soup.select_one('dt:-soup-contains("ショルダー") + dd'):
        #     shoulder_strap_cm = page_soup.select_one('dt:-soup-contains("ショルダー") + dd')
        #     if '-' in shoulder_strap_cm.text or '~' in shoulder_strap_cm.text:
        #         shoulder_strap_cm = ""
        #         print('='*30)
        #         print('Success')
        #     else:
        #         shoulder_strap_cm = float(shoulder_strap_cm.text.replace('約',' ').replace('c', '').replace('m', '')) if shoulder_strap_cm else ''
        # else:
        #     shoulder_strap_cm = ''

        # shop_address = shop_address.text if shop_address else None
        #- else Noneの代わりに ''で入力可能!!
        #! 測定方法を間違えていた inch
        body_width_inch = round(body_width_cm / 2.54, 2) if body_width_cm else ''

        shoulder_inch = round(shoulder_cm / 2.54, 2) if shoulder_cm else ''

        sleeve_length_inch = round(sleeve_length_cm / 2.54, 2) if sleeve_length_cm else ''

        length_inch = round(length_cm / 2.54, 2) if length_cm else ''

        #- 関数_3を使用
        actual_size_cm = to_actual_size(body_width_cm, shoulder_cm, sleeve_length_cm, length_cm,'cm')

        actual_size_inch = to_actual_size(body_width_inch, shoulder_inch, sleeve_length_inch, length_inch,'inch')

        d_list.append({
            '*Action(SiteID=US|Country=JP|Currency=USD|Version=745|CC=UTF-8)': 'Add',
            #- 20220708 - 服のサイズ
            # 'DL-eb-category': 'bag_m:52357 / bag_w:169291',
            'X-tag-size': tag_size,
            #- 要結果確認_excel上で'003のようにしないと00が反映されないため_20220223
            'X-number': "'"+str(iterated_num).zfill(3),
            '*Title': '★①-titleで  C:Type  /  C:Style を確認の上、入力する！例: Jacket or Coat, 2つとも同じ値で良い - ②[ConditionDescription] 毛100％等、毛はhairと訳されるので、woolに置き換えること! - ③Size の中身が "-" のものは変更すること upload失敗するため - ④urlが重複していないか要確認!!',
            #! NEW 20220211 最終的に*Titleにコピペするための
            #! ご操作を防ぐために別にタイトルを記載
            # 'U-title': item_title,
            'U-title': delete_brackets(item_title),
            #! NEW 20220210 タイトルに付属させる用途
            'X-category':category_whole,
            #- 20220708
            # 'Ref-title_category_combined':'^^=E2&" "&F2',
            'Ref-title_category_combined':'^^=^^E2',
            'Ref-translation':'^^=if(G2="","",googletranslate(G2,"ja","en"))',
            #- タイトルにモデル名と識別番号を付与した_20220219
            'URef-title-completed':'^^=H2&" "&AF2&" "&"from Japan"&" "&C2',
            'Ref-len()': '^^=len(D2)',
            'Ref-material_translation': '^^=if(AG2="","",googletranslate(AG2,"ja","en"))',
            # 'URef-size_for_ConditionDescription': '^^=M2&" | "&N2&" | "&" Model: "&AF2&" | "&"Material: "&K2&" | "&"Shipping from Japan. For more details, please check the pictures carefully and judge the condition.',
            #- update - 20220621 商品の状態を追加した AA2
            'URef-size_for_ConditionDescription': '^^=AA2&" | "&"Tag size: "&B2&" | "&M2&" | "&N2&" | "&" Model: "&AF2&" | "&"Material: "&K2&" | "&"Shipping from Japan. Note: All you receive is what you see in the pictures(a hanger or a stand is not included.). For more details, please check the pictures carefully and judge the condition."',
            'X-actual_size_inch': actual_size_inch,
            'X-actual_size_cm': actual_size_cm,
            'CustomLabel' : 'Please fill in like this. py_Supplier-BRAND-Category_Gender-DATE-Number',
            '*Description' : '',
            #¥ New 20220211
            #- ConditionDescription にはサイズや型番、状態が入る
            'ConditionDescription':'★★★★★★★★★',
            '*StartPrice' : '',
            'MinimumBestOfferPrice': '"¥=R2',
            'url': item_url,
            'costPrice + s/p': item_price + 550,
            # 'combined':'=c{num}+h{num}'.format(num = i+2),
            'ConditionID' : '3000',
            'PicURL':zoom_pictures,
            # 'C:Brand': brand_name,
            'C:Brand': delete_brackets(brand_name),
            's/p': shipping_cost,
            'X-gender': gender,
            # 'X-condition': condition,
            #- condition_rankで商品の状態を明確にした 20220621 update
            'X-condition': condition_rank,
            #- 不要なので削除及び下記に修正 x-conditionと相違ないか確認するため - 20220621
            # 'X-accessories': accessories,
            'X-condition_info': condition,
            'X-item_detail': item_detail,
            'X-staff_comment': staff_comment,
            'X-color': color,
            'X-model': model,
            'X-material': material,
            'X-manufactured': manufactured,
            'X-remark': remark,
            'X-condition_detail': condition_detail,
            # 'X-all_sizes': store_all_sizes,
        })

    print('============終わり===============')
    sleep(1)

#- update - 20220621
save_location = control_path_class.neutralize_path_str('csv_folder/csv_for_concat/ebay')

#- update - 日本語を含まない英語のみのブランド名を取得 - 20220621
brand_name_revised = delete_brackets(brand_name)
file_name = f'{save_location}/trf-{brand_name_revised}-bag-{category_gender}-'

print('\n','保存場所: ',file_name)

df = pd.DataFrame(d_list)

if is_category_men == True:
    df_add = pd.read_excel('e_category_no/jacket_coat_vest_men.xlsx')
else:
    #! 要変更!!
    df_add = pd.read_excel('e_category_no/bag_women.xlsx')

df_complete = pd.concat([df,df_add],axis=1)

df_complete.to_excel(file_name + today + '.xlsx', index=False)

driver.close()

now = datetime.datetime.now(JST)
now = now.strftime('%Y-%m-%d %H:%M:%S')
print('\n','★★★終了時刻: ',now, '\n')

folder_dir = control_path_class.neutralize_path_str(save_location)
print('folder_dir: ', folder_dir)

#- プログラム終了後、保存先フォルダーを開く  https://tinyurl.com/25ojjpdl
module_01.open_folder(folder_dir)