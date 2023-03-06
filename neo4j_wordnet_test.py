from neo4j import GraphDatabase
import neo4j_write_synset as write_neo4j
from nltk.corpus import wordnet as wn

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

# 深さ優先探索
def search_DFS(root):
    children = get_hyponyms(root)
    if children == []:
        return
    else:
        for child in children:
            session.write_transaction(write_neo4j.hyponym_link, str(root), str(child))
            search_DFS(child)

def get_hyponyms(synset):
    hyponyms = synset.hyponyms()
    return hyponyms

def hyponym_link(tx, name, friend_name=None):
    if not friend_name:
        return tx.run('CREATE (p:Synset {name: $name}) RETURN p', name=name)
    tx.run('MATCH (p:Synset {name: $name})'
           'CREATE (p)-[:Hyponym]->(:Synset {name: $friend_name})',
            name=name, friend_name=friend_name)
    tx.run('MATCH (p:Synset {name: $name})'
           'CREATE (p)<-[:Hyperym]-(:Synset {name: $friend_name})',
            name=name, friend_name=friend_name)


with driver.session() as session:
    session.write_transaction(write_neo4j.hypernym_link, str(wn.synsets('entity', lang='eng')[0]))
    search_DFS(wn.synsets('entity', lang='eng')[0])