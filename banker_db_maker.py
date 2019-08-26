""" Design Doc:
    create db with following fields:
        [account, realm, name, bank_num, bank_gold, trade_timestamp, trade_confirmation]
"""

from farmer_data import *


BANKERS_DB = 'output_files/bankers.sqlite3'
wd = WowData()
wd.update_bankers()

def create_table():
    conn = sqlite3.connect(BANKERS_DB)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS bankers (
                 id INTEGER PRIMARY KEY,
                 name TEXT,
                 realm TEXT,
                 account INTEGER,
                 bank_num INTEGER,
                 bank_gold INTEGER,
                 trade_timestamp INTEGER,
                 trade_confirmation INTEGER,
                 UNIQUE(name, realm))""")
    c.execute("""CREATE TABLE IF NOT EXISTS realms (
                 id INTEGER PRIMARY KEY,
                 realm TEXT,
                 code TEXT,
                 price_per_mil INTEGER,
                 UNIQUE(realm))""")
    lower_price = ['Hyjal', 'Blackmoore', 'Blackrock', 'Antonidas', 'Archimonde', 'Blackhand']
    for realm in wd.realms.values():
        c.execute("""INSERT OR IGNORE INTO realms
                     (realm, code, price_per_mil)
                     VALUES(?, ?, ?)""",
                     (realm.name, realm.code, 41 if realm.name in lower_price else 43))
    conn.commit()
    conn.close()

def update_table():
    create_table()
    conn = sqlite3.connect(BANKERS_DB)
    c = conn.cursor()
    for banker in wd.bankers.values():
        if (banker.bank_gold and banker.bank_gold > 100):
            values = (banker.name,
                      banker.realm,
                      banker.account,
                      banker.bank_number,
                      banker.bank_gold)
            
            c.execute("""INSERT OR IGNORE INTO bankers
                         (name, realm, account, bank_num, bank_gold)
                         VALUES(?, ?, ?, ?, ?)""", values)

    conn.commit()
    conn.close()

if __name__ == '__main__':
    update_table()