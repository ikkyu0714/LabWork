from neo4j import GraphDatabase

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

def add_friend(tx, name, friend_name=None):
    if not friend_name:
        return tx.run('CREATE (p:Person {name: $name}) RETURN p', name=name)
    return tx.run('MATCH (p:Person {name: $name})'
                  'CREATE (p)-[:FRIEND]->(:Person {name: $friend_name})',
                  name=name, friend_name=friend_name)

def print_friend(tx, name):
    for record in tx.run('MATCH (p {name: $name})-[:FRIEND]->(yourfriends)'
                         'RETURN p,yourfriends', name=name):
        print(record)

with driver.session() as session:
    session.write_transaction(add_friend, 'MasterU')
    for f in ['Mark', 'Kent']:
        session.write_transaction(add_friend, 'MasterU', f)
    #session.read_transaction(print_friend, 'MasterU')

"""from neo4j import GraphDatabase

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

def clear_db(tx):
    tx.run('MATCH (n) DETACH DELETE n')

with driver.session() as session:
    session.write_transaction(clear_db)"""