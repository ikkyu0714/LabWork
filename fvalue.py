import openpyxl
import numpy as np
from sklearn.metrics import accuracy_score, f1_score, precision_score, recall_score

def threshold(value):
    if value >= 0.5:
        return 1
    elif value < 0.5:
        return 0

EXCEL_PATH = '../CDD/new_datasetのコピー.xlsx'
book = openpyxl.load_workbook(EXCEL_PATH)
sheet = book['Sheet1']

y_true = []
ori_pred = []
k_pred = []
iso_pred = []
db_pred = []
for line in sheet.iter_rows(min_row=2, min_col=2):
    list = []
    for item in line:
        list.append(item.value)
    if list[3] == '文化差なし':
        y_true.append(1)
    else:
        y_true.append(0)
    ori_pred.append(threshold(list[4]))
    k_pred.append(threshold(list[5]))
    iso_pred.append(threshold(list[6]))
    db_pred.append(threshold(list[7]))

# 保存
book.save(EXCEL_PATH)
# 終了
book.close()

print('------------- 異常値検出なし ----------------')
print('正確性:{}, 適合率:{}, 再現率:{}, F値:{}'.format(accuracy_score(y_true, ori_pred), precision_score(y_true, ori_pred), recall_score(y_true, ori_pred), f1_score(y_true, ori_pred)))
print('-------------    K-means   ----------------')
print('正確性:{}, 適合率:{}, 再現率:{}, F値:{}'.format(accuracy_score(y_true, k_pred), precision_score(y_true, k_pred), recall_score(y_true, k_pred), f1_score(y_true, k_pred)))
print('-------------    DBSCAN    ----------------')
print('正確性:{}, 適合率:{}, 再現率:{}, F値:{}'.format(accuracy_score(y_true, db_pred), precision_score(y_true, db_pred), recall_score(y_true, db_pred), f1_score(y_true, db_pred)))
print('-------------   Isolation  ----------------')
print('正確性:{}, 適合率:{}, 再現率:{}, F値:{}'.format(accuracy_score(y_true, iso_pred), precision_score(y_true, iso_pred), recall_score(y_true, iso_pred), f1_score(y_true, iso_pred)))

#print(y_true)