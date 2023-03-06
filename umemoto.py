from sklearn.svm import SVC
from sklearn.preprocessing import StandardScaler
from vgg16 import features
from gram import list_int
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
#scikit-learn内にあるワインの品質判定用データセットをwineという変数に代入。
features_float = []
for feature in features:
    features_float.append(feature.astype(float))

gram=list_int
#Xにワインの特徴量を代入。特徴量は、ここではアルコール濃度や色などwineの属性データのこと。
X=features_float
#yにワインの目的変数を代入。目的変数は、ここでは専門家によるワインの品質評価結果のこと。
y=gram
print(y)
#トレーニングデータとテストデータに分割。
#トレーニングデータで学習を行い、テストデータでAIが実際に使えるかの精度検証。テストデータは全体の3割に設定。
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=0)
#標準化を実行するためのStandardScalerをインスタンス化（実体化）。fitで訓練データを標準化する際の準備。
scaler = StandardScaler()
scaler.fit(X_train)
#サポートベクターマシンをインスタンス化（実体化）。 回帰の場合はSVR()。
clf = SVC()
#transformで標準化を実際に行い、そのデータに対してサポートベクターマシンによる学習を行う
clf.fit(scaler.transform(X_train), y_train)
#テストデータにも標準化を実行し、predict(予測)を行う
y_pred = clf.predict(scaler.transform(X_test))
#正解率の算出。予測データと正解データを比較してAIの精度検証を行う。
accuracy_score(y_test,y_pred)