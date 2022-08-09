#
#to 参考 csv結合: https://coffee-blue-mountain.com/python-csv-comb2/#toc3

#to 参考 時間 : https://atmarkit.itmedia.co.jp/ait/articles/2111/09/news015.html

from site import addpackage
import pandas as pd
import glob
import datetime
#- シート名の有無によってconcatするファイルを指定するため読み込む
import openpyxl

#- 20220530
import os
from pathlib import Path

t_delta = datetime.timedelta(hours=9)
JST = datetime.timezone(t_delta, 'JST')
now = datetime.datetime.now(JST)
# print(repr(now))

today = now.strftime('%Y%m%d')

d = now.strftime('%Y%m%d%H%M%S')
# print(d)  # 20211104173728 秒まで

d = now.strftime('%Y%m%d%H%M')
# print(d)  # 202111041737 分まで

#- windows変更 \を/に変更した- 20220529
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

#- 上記だと、作成したmain_data~.xlsxのファイルまでconcatしてしまうため、下記の上に修正
#-ファイル名の文字列によって、concatするファイルを選別する-20220428
# https://teratail.com/questions/142775
excel_files = []

#- 20220530
file_location_obj = neutralize_path_obj('csv_folder/csv_for_concat/ebay')
# file_location_obj = 'csv_folder/csv_for_concat/ebay'
glob_files = list(file_location_obj.glob('*.xlsx'))
print(type(glob_files))
for excel_file_name in glob_files:
    #! listにしていても、TypeError: argument of type 'WindowsPath' is not iterable が発生するので、その後の要素をstr()で囲んだところ、解消した。
    excel_file_name = str(excel_file_name)
    if 'main' in excel_file_name:
    #* continueしても後続のifは実行される - if,elseとは違う
        continue
    if 'trf' in excel_file_name:
        excel_files.append(excel_file_name)
    if '2nd' in excel_file_name:
        excel_files.append(excel_file_name)

print('worked')

#読み込むファイルのリストを表示
# for excel_file in excel_files:
#     print('excel_file: ', excel_file)

# concatするファイル一覧を表示
print('\n', 'concatするファイル一覧: ',excel_files,'/ concatするファイル数: ',len(excel_files),'\n')

#csvファイルの中身を追加していくリストを用意
data_list = []

#読み込むファイルのリストを走査
for file in excel_files:
    #- Excelのシート名をすべて取得し、シート名によって読み込むシートを指定する-20220527
    # https://tinyurl.com/2m3rl3qo
    # ブックを取得
    book = openpyxl.load_workbook(file)
    # シートを取得 シート名がリストで取得される 例:['main', 'reference']
    sheets = book.sheetnames
    #^ リストの要素に'reference'シートがあった場合、'reference'シートを読み込む
    if 'reference' in sheets:
        data_list.append(pd.read_excel(file,sheet_name='reference'))
    else:
        data_list.append(pd.read_excel(file))

#リストを全て行方向に結合

# print('data_list',data_list)

df = pd.concat(data_list)
# excel_title  = 'main_data_'+ d+'.xlsx'
excel_title  = 'main_data_'+ today +'.xlsx'
print(excel_title)

#! bag_women.xlsxのFXの残りの部分が縦にも横にも長過ぎたためか、普通にconcatをしたら、最初の二行だけ出力されて残りが出てこなかったと思ったが、FXの行が長すぎただけで、恐らくずっと下のほうに値は存在していた。ブランド名でフィルターかけてみて今後は確認してみよう!
#- 要停止
# df_add = pd.read_excel('e_category_no/bag_women.xlsx')

#¥ pandas.errors.InvalidIndexError: Reindexing only valid with uniquely valued Index objects
#to 参考:   https://tinyurl.com/ya3uv8gk
df.reset_index(inplace=True, drop=True)

path = neutralize_path_str('csv_folder/csv_for_concat/ebay')
df.to_excel(f'{path}//{excel_title}',index=False)

print('\n', '保存場所: ', f'{path}/{excel_title}' ,'\n')

print()

#! #! 一時停止 => concatがうまく行かなかった場合用
# df_complete.to_excel("csv_folder/csv_for_concat/"+excel_title,index=False)
