# Wordnetをインポート
from nltk.corpus import wordnet as wn
# Networkxをインポート
import networkx as nx
import matplotlib
import matplotlib.pyplot as plt
import sys

class All_Synstes_Get():
    def __init__(self):
        self.G = nx.Graph()
        self.search_list = []
        #children = self.get_hyponyms(root)

    def search_DFS(self, root):
        children = self.get_hyponyms(root)
        if children == []:
            return
        else:
            for child in children:
                self.G.add_edge(root, child)
                self.search_DFS(child)

    def output_graph(self):
        values = [1.0 for node in self.G.nodes()]
        edge_colours = ['black' for edge in self.G.edges()]
        black_edges = [edge for edge in self.G.edges()]
        pos = nx.spring_layout(self.G)
        nx.draw_networkx_nodes(self.G, pos, cmap=plt.get_cmap('Reds'), node_color = values, node_size=4800)
        nx.draw_networkx_edges(self.G, pos, edgelist=black_edges, arrows=True)

        nx.draw_networkx_labels(self.G,pos,font_size=12,font_family='TakaoExMincho')

        plt.axis('off')
        plt.show()

    def search_BFS(self, root):
        children = self.get_hyponyms(root)
        for child in children:
            self.G.add_edge(root, child)
            self.search_list.append(child)
        if self.search_list != []:
            self.search(self.search_list.pop(0))

    def start_search_all_hyponyms(self):
        self.search_hyponyms(wn.synsets('entity'))

    def search_hyponyms(self, synsets):
        for synset in synsets:
            print(synset)
            hyponyms = self.get_hyponyms(synset)
            if hyponyms == []:
                return None
            else:
                return self.search_hyponyms(hyponyms)

    def get_hyponyms(self, synset):
        hyponyms = synset.hyponyms()
        return hyponyms

class Test():
    def __init__(self, oya):
        self.G = nx.Graph()
        self.root = 'START ->'
        children = self.get_hyponyms(oya)
        for child in children:
            self.walks(oya, child)

        values = [1.0 for node in self.G.nodes()]
        edge_colours = ['black' for edge in self.G.edges()]
        black_edges = [edge for edge in self.G.edges()]
        pos = nx.spring_layout(self.G)
        nx.draw_networkx_nodes(self.G, pos, cmap=plt.get_cmap('Reds'), node_color = values, node_size=4800)
        nx.draw_networkx_edges(self.G, pos, edgelist=black_edges, arrows=True)

        nx.draw_networkx_labels(self.G,pos,font_size=12,font_family='TakaoExMincho')

        plt.axis('off')
        #plt.show()
        print(len(self.G.nodes()))
        print("fin")
        
    def walks(self, oya, child):
        self.G.add_edge(oya, child)
        self.root = self.root + str(child) + ' -> '
        oya = child
        children = self.get_hyponyms(oya)
        if len(children) != 0:
            for child in children:
                return self.walks(oya, child)

    def get_hyponyms(self, synset):
        hyponyms = synset.hyponyms()
        print(hyponyms)
        return hyponyms

all_synsets_get = All_Synstes_Get()
all_synsets_get.search_DFS(wn.synsets('entity')[0])
print(len(all_synsets_get.G.nodes()))
all_synsets_get.output_graph()

"""
test = Test(wn.synsets('entity')[0])
print(test.root)
#print(wn.synsets('entity'))"""