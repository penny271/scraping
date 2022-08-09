#! reCAPTCHA V2 の "私はロボットではありません"と認証が入るため使用不可能
#! https://www.2ndstreet.jp/user/login ループの2回目で認証が入っている模様
#! seleniumを使って、認証させてログインしてから通常通りにbeautifulSoupでスクレイピングが可能か？
from time import sleep

import requests
from bs4 import BeautifulSoup
import pandas as pd

ua = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_6) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Safari/605.1.15'

headers = {
  'User-Agent' : ua
}

d_list = []
base_url ='https://www.2ndstreet.jp/search?category=950001&keyword=nike&sortBy=cost-low&page={}'

for i in range(1):
  print('='*30, i, '='*30)
  url = base_url.format(i+1)

  sleep(2)

  r = requests.get(url, headers = headers, timeout=3)
  r.raise_for_status()

  soup = BeautifulSoup(r.content, 'lxml')
  page_urls = soup.select('ul[class="itemList"] > li[class="js-favorite"] > a')

  #! At the second loop, error occurs
  for i, page_url in enumerate(page_urls):
    page_url = 'https://www.2ndstreet.jp'+page_url.get('href')
    print(page_url)
    print('=====Loop=====',i)
    sleep(2)

    page_req = requests.get(url, headers = headers, timeout=3)
    page_req.raise_for_status()

    page_soup = BeautifulSoup(page_req.content, 'lxml')

    zoom_pictures = []
    pictures = page_soup.select('div[class="goods_popname"] > ul > li')

    for i, picture in enumerate(pictures):
      picture = picture.select_one('img').get('data-src')
      zoom_pictures.append(picture)

    brand_name = page_soup.select_one('p.blandName').text.strip()
    print(brand_name)

    #!/user/login/goods/detail/goodsId/2330000369738/shopsId/30943
