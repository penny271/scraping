#! seleniumなどモジュールがインストールされていないというエラーが出たら、vscodeをanacondaのenv環境で開けていないため、ANACONDA.NAVIGATORを起動させてから、ターミナルを起動する!! _2022/02/07

from time import sleep

import requests
from bs4 import BeautifulSoup
import pandas as pd

#! user-agentの偽装 - 20220215
ua = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15'

headers = {
  'User-Agent' : ua
}

d_list = []
#- 取得したいブランド名を元のサイトのサーチボックスで検索後、そのurlのページ数の部分を{}に書き換え、base_urlとする
base_url ='https://www.yourshoppingmap.com/brand/295-moorer?page={}#boutiques'

# 'https://www.yourshoppingmap.com/brand/295-moorer?page=2#boutiques'

#- 要rangeの中の数をpage数に合わせて変更
for i in range(9):
  print('='*30, i, '='*30)
  url = base_url.format(i+1)

  sleep(3)

  # url = 'https://www.yourshoppingmap.com/brand/295-moorer'

  r = requests.get(url, headers = headers, timeout=3)
  r.raise_for_status()

  print(r.text)

  soup = BeautifulSoup(r.content, 'lxml')
  # page_urls = soup.select('a:-soup-contains("企業ページ")')
  page_urls = soup.select('div[class="title desktop-only"]>a')

  # print(page_urls)

  for page_url in page_urls:
    # page_url = 'https://next.rikunabi.com' + page_url.get('href')

    page_url = page_url.get('href')

    # print(page_url)

    sleep(3)

    page_r = requests.get(page_url, timeout=3)
    page_r.raise_for_status()

    page_soup = BeautifulSoup(page_r.content, 'lxml')

    #! 複数のセレクタで対象を絞り込む(区切り文字なし) 対象の2つのクラスを持つもののみ(間にスペースを入れない)を指定 strip()で空白削除
    shop_name = page_soup.select_one('.title.single').text.strip()
    shop_address = page_soup.select_one('address')
    #! 三項演算子
    shop_address = shop_address.text if shop_address else None
    #! 2つのクラスを含む要素の隣の要素aタグの要素
    shop_url = page_soup.select_one('.icon.icon-world+span > a')
    shop_url = shop_url.get('href') if shop_url else None
    #! 2つのクラスを含む要素の隣の要素aタグの要素
    shop_email = page_soup.select_one('.icon.icon-mail+span > a')
    shop_email = shop_email.text if shop_email else None
    #! 2つのクラスを含む要素の隣の要素の中のaタグの要素
    shop_phone = page_soup.select_one('.icon.icon-phone+span > a')
    shop_phone = shop_phone.text if shop_phone else None

    #- 複数ある.accordion-content要素 をリストで取得
    brands = page_soup.select('.accordion-content')
    # print(brands)
    #- blank_box = []
    blank_box_men = [] # for Men
    blank_box_women = [] # for Women
    blank_box_shoes = [] # shoes and accessory
    blank_box_fragrances = [] # fragrance

    #- 複数ある.accordion-content要素 をリストから一つずつ取得
    #! enumerate()は不要だった
    for i,brand in enumerate(brands):

      # if i ==0:
      #! spanタグでテキスト Man を含む要素を取得
      if page_soup.select_one('span:-soup-contains("Man")'):
        brand_men = brand.select('a')
        for brand_man in brand_men:
          blank_box_men.append(brand_man.text)
          print(blank_box_men[-1])
      if page_soup.select_one('span:-soup-contains("Woman")'):
        brand_women = brand.select('a')
        for brand_woman in brand_women:
          blank_box_women.append(brand_woman.text)
          # print(blank_box_women[-1])
      if page_soup.select_one('span:-soup-contains("Shoes and accessories")'):
        brand_shoes = brand.select('a')
        for brand_shoe in brand_shoes:
          blank_box_shoes.append(brand_shoe.text)
          # print(blank_box_shoes[-1])
      if page_soup.select_one('span:-soup-contains("Fragances")'):
        brand_fragrances = brand.select('a')
        for brand_fragrance in brand_fragrances:
          blank_box_fragrances.append(brand_fragrance.text)
          # print(blank_box_fragrances[-1])
      else:
        break

    #- url_in_tag = page_soup.select_one('.rnn-col-11:last-of-type a')
    # shop_url = url_in_tag.get('href') if url_in_tag else None

    # shop_tel = page_soup.select_one('.class=icon icon-phone > span > a.get("href")').text

    d_list.append({
      'Name': shop_name,
      'URL': shop_url,
      'Address': shop_address,
      'Tel': shop_phone,
      'Email':shop_email,
      'Brand - Men' : blank_box_men,
      'Brand - Women' : blank_box_women,
      'Brand - Shoes' : blank_box_shoes,
      'Brand - Fragrances' : blank_box_fragrances,
    })
    # print(d_list[-1])

df = pd.DataFrame(d_list)
df.to_csv('shop_list_MOORER_20220204.csv', index=None, encoding='utf-8-sig')
