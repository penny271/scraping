#^ 正規表現を使うライブラリ
import re

from bs4 import BeautifulSoup
import requests

url = 'https://www.python.org/'
r = requests.get(url)

soup = BeautifulSoup(r.content, 'lxml')
post = soup.find('div', class_='blog-widget')

#* 完全一致したテキストを見つける
print(post.find(text='Latest News').parent)
print(soup.find(href='/about/'))

#^ 部分一致した要素をテキストを取得
#* text= だけだと文字列だけ取得となる
for text in soup.find_all(text=re.compile('Python')):
  # print(text)
  print(text.parent)

#- Pythonという文字列が含まれているaタグだけ取得する
for text in soup.find_all('a',text=re.compile('Python')):
  # print(text)
  print(text)

print('===========================================')

#- Pythonという文字が含まれているurlを含むaタグだけ取得する
for text in soup.find_all('a',href=re.compile('docs.python.org')):
  # print(text)
  print(text)
