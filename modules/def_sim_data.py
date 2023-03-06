# Excel_Dataクラスを定義
class Excel_Data:
    def __init__(self, id, japanese, english, answer, similarity, hypernym):
        self.id = id
        self.japanese = japanese
        self.english = english
        self.answer = answer
        self.similarity = similarity
        self.hypernym = hypernym
