import openpyxl
import numpy
import matplotlib.pyplot as plt
import pandas

EXCEL_PATH = 'dataset0907.xlsx'
book = openpyxl.load_workbook(EXCEL_PATH)
sheet = book['Sheet1']

noculture_simi = []
culture_simi = []
similarity = []
fig = plt.figure(figsize = (5,5), facecolor='white')
plt.title("Similarity Histogram", fontsize=15)  # (3) タイトル
plt.xlabel("Similarity", fontsize=15)            # (4) x軸ラベル
plt.ylabel("Frequency", fontsize=15)      # (5) y軸ラベル

for line in sheet.iter_rows(min_row=3):
    values = []
    for item in line:
        values.append(item.value)
    #文化差のラベルごと
    #noculture_simi.append(values[4]) if values[3] == '文化差なし' else culture_simi.append(values[4])
    similarity.append(values[4])

#文化差のラベルごとに分ける
#plt.hist([noculture_simi, culture_simi], bins=21, range=(0.0, 1.0), ec='black', label=['NoCulture', 'Culture'])
#plt.legend(loc='upper left')
plt.hist(similarity, bins=21, range=(0.0,1.0), ec='black')
plt.show()
# 保存
book.save(EXCEL_PATH)
# 終了
book.close()