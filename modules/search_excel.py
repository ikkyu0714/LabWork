class Search_Excel():
    def __init__(self):
        self.count = 0 # 出現回数を数えるカウンター
        self.list =[] # データを格納するリスト
        #self.correct = [] # 検索に引っ掛かった概念のidをリストに保存
        #self.error = [] # 判定失敗した概念のidをリストに保存
        #self.target = [] # 対象物 判定失敗かつ検索wordを含む概念のidをリストに保存
        #self.subtarget = [] # 対象物 判定成功かつ検索wordを含む概念のidをリストに保存

    def add(self, data):
        self.list.append(data) # リストに要素を追加
