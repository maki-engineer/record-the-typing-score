import os
import glob

# TWellJR.exeファイル探索
file_search = [file_name for file_name in glob.glob("TWellJR.exe")]

if not file_search:
  print("\n　TWellJR.exeファイルが見つかりませんでした...\n　タイピング記録.exeは、TWellJR.exeファイルが置いてある場所に置いてください！\n")
  os.system('PAUSE')
  exit()

# JR全履歴フォルダ探索
directory_search = [directory_name for directory_name in glob.glob("JR全履歴")]

# 見つからなかったら作っちゃって終わり
if not directory_search:
  os.mkdir("JR全履歴")
  print("　タイプウェルをされていないため、記録しませんでした。")
  os.system('PAUSE')
  exit()

# 本日分のテキストファイル名作成
import datetime

now         = datetime.datetime.now()
format_year  = str(now.year % 100)
format_month = now.strftime("%m")
format_date  = now.strftime("%d")
text_file    = "./" + directory_search[0] + "/" + format_year + format_month + format_date + "t.txt"

# 本日分のテキストファイル探索
today_score_search = [file_name for file_name in glob.glob(text_file)]

if not today_score_search:
  print("　タイプウェルをされていないため、記録しませんでした。")
  os.system('PAUSE')
  exit()

# 本日分のテキストファイルの記録の中から最高記録を探索、その最高記録のスコアを計算
import math

# スコア計算の関数
def score_calc(per_sec, miss):
  wpm               = per_sec * 60
  correct_percent   = 400 / (400 + miss)
  correct_percent **= 3
  result            = math.floor(wpm * correct_percent)

  return [math.floor(wpm), result]

# ここで最高記録を探索する処理
import time as wait

text_file_detail = open(today_score_search[0], 'r')
file_insides     = text_file_detail.readlines()

high_scores = []
high_score  = 0

for file_inside in file_insides:
  data    = file_inside.split(",")
  time    = float(data[3].strip())
  miss    = int(data[13].strip())
  per_sec = round(400 / time, 2)
  scores  = score_calc(per_sec, miss)

  if high_score < scores[1]:
    high_score  = scores[1]
    high_scores = [scores[0], miss, scores[1]]

text_file_detail.close()

print("\n　エクセルに本日の最高スコアを記録しています...\n")
wait.sleep(5)

# ここでエクセルに記録
import openpyxl as px

wb          = px.load_workbook(filename = ".//タイピングスコア.xlsx")
sh          = wb['タイピング']
month, date = now.month, now.day
today       = str(month) + '月' + str(date) + '日'
row_index   = 3

# どこに記録すればいいのか探索
while True:
  if sh['B' + str(row_index)].value:
    row_index += 1
  else:
    # row_index の行に記録
    sh['B' + str(row_index)].value = today
    sh['C' + str(row_index)].value = high_scores[0]
    sh['D' + str(row_index)].value = high_scores[1]
    sh['E' + str(row_index)].value = high_scores[2]
    break

# 保存
wb.save(".//タイピングスコア.xlsx")

print("\n　記録が完了しました！\n")

os.system('PAUSE')
