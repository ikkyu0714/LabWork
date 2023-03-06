from neo4j import GraphDatabase

driver = GraphDatabase.driver('bolt://localhost:7687', auth=('neo4j', 'ikkyu2419'))

def clear_db(tx):
    tx.run('MATCH (n) DETACH DELETE n')

with driver.session() as session:
    session.write_transaction(clear_db)