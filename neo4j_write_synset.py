from neo4j import GraphDatabase

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

def hyponym_link(tx, name, friend_name=None):
    if not friend_name:
        return tx.run('CREATE (p:Synset {name: $name}) RETURN p', name=name)
    return tx.run('MATCH (p:Synset {name: $name})'
                  'CREATE (p)-[:Hyponym]->(:Synset {name: $friend_name})',
                  name=name, friend_name=friend_name)

def hypernym_link(tx, name, friend_name=None):
    if not friend_name:
        return tx.run('CREATE (p:Synset {name: $name}) RETURN p', name=name)
    return tx.run('MATCH (p:Synset {name: $name})'
                  'CREATE (p)-[:Hypernym]->(:Synset {name: $friend_name})',
                  name=name, friend_name=friend_name)

def print_friend(tx, name):
    for record in tx.run('MATCH (p {name: $name})-[:FRIEND]->(yourfriends)'
                         'RETURN p,yourfriends', name=name):
        print(record)

if __name__ == '__main__':
    with driver.session() as session:
        session.write_transaction(hyponym_link, 'Burdock')
        for f in ['Great Burdock', 'Common Burdock']:
            session.write_transaction(hyponym_link, 'Burdock', f)
            session.write_transaction(hypernym_link, f, 'Burdock')
            #session.write_transaction(hypernym_connect, f, 'Burdock')
        #session.read_transaction(print_friend, 'MasterU')

"""from neo4j import GraphDatabase

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

def clear_db(tx):
    tx.run('MATCH (n) DETACH DELETE n')

with driver.session() as session:
    session.write_transaction(clear_db)"""