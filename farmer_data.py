import json
import os
import sqlite3
import time

from slpp import slpp as lua
from openpyxl import load_workbook
from openpyxl.styles import Font

from settings import *


class AccData:
    def __init__(self, path):
        def table_to_dict():
            """Decodes lua SavedVariables table as dict."""
            if not os.path.exists(path):
                return {}
            data = {}
            with open(path, encoding="ISO-8859-1") as file:
                if file:
                    data = lua.decode(file.read().replace('Multiboxer_DataDB = ', ''))
                    return data

        self.path = path
        data = table_to_dict()
        self.characters = data['charData']
        self.realms = data['realmData']


class Farmer:
    def __init__(self, data):
        self.name = data[0]
        self.realm = data[1]
        self.account = data[2]
        self.char_class = data[3]
        self.intro_completed = None
        self.professions = {}
        self.recipe_ranks = {}

    def update_info(self, data):
        self.intro_completed = data.get('introCompleted', None)
        self.professions = {}
        for profession_id, rating in data.get('professions', None).items():
            if type(rating) is int:
                self.professions[profession_id] = rating
        self.recipe_ranks = {}
        for recipe, rank in data.get('recipeRanks', None).items():
            self.recipe_ranks[recipe] = rank


class FarmerData:
    def __init__(self):
        def farmer_objects_dict():
            conn = sqlite3.connect(CHAR_DB)
            c = conn.cursor()
            c.execute("""SELECT name, realm, account, class FROM char_db
                    WHERE role=? AND (type!=? or type IS NULL)""",
                    ('farmer', 'inactive'))

            farmers = {}
            for row in c.fetchall():
                name_realm = '-'.join((row[0], row[1]))
                farmers[name_realm] = Farmer(row)
            return farmers

        def account_objects_dict():
            accounts = {}
            for acc_number, path in LUA_PATHS.items():
                accounts[int(acc_number)] = AccData(path)
            return accounts  

        def create_output_db():
            """write doc"""
            
            conn = sqlite3.connect(FARMERS_DB)
            c = conn.cursor()
            c.execute("""CREATE TABLE IF NOT EXISTS farmers (
                name TEXT NOT NULL,
                realm TEXT NOT NULL,
                account INTEGER NOT NULL,
                class TEXT,
                intro_complete INTEGER,
                Herbalism INTEGER,
                Mining INTEGER,
                "Zin'anthid" INTEGER,
                "Osmenite Deposit" INTEGER,
                "Osmenite Seam" INTEGER)""")

        self.farmers = farmer_objects_dict()
        self.accounts = account_objects_dict()
        create_output_db()

    def update_farmers(self):
        """write doc"""
        for acc_num in self.accounts.keys():
            char_list = self.accounts[acc_num].characters
            for full_name, char_data in char_list.items():
                self.farmers[full_name].update_info(char_data)

    def write_farmers_db(self):
        """write doc"""
        conn = sqlite3.connect(FARMERS_DB)
        c = conn.cursor()
        c.execute("DELETE FROM farmers")

        for full_name, farmer in self.farmers.items():
            values = (farmer.name,
                      farmer.realm,
                      farmer.account,
                      farmer.char_class,
                      1 if farmer.intro_completed else 0,
                      farmer.professions.get(2549, None),
                      farmer.professions.get(2565, None),
                      farmer.recipe_ranks.get("Zin'anthid", None),
                      farmer.recipe_ranks.get('Osmenite Deposit', None),
                      farmer.recipe_ranks.get('Osmenite Seam', None),)
            c.execute("""INSERT INTO farmers
                    (name, realm, account, class, intro_complete, Herbalism, Mining,
                    "Zin'anthid", "Osmenite Deposit", "Osmenite Seam")
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    values)
        conn.commit()
        conn.close()

    def write_excel_table(self):
        wb = load_workbook(EXCEL_PATH, read_only=False, keep_vba=True)
        ws = wb["Farming"]
        selection = ws["B32:K57"]

        conn = sqlite3.connect(FARMERS_DB)
        c = conn.cursor()

        realm_list = ['Kazzak', 'Twisting Nether', 'Silvermoon', 'Draenor', 'Tarren Mill', 'Ragnaros', 'Ravencrest', 'Argent Dawn', 'Sylvanas', 'Frostmane', 'Burning Blade', 'Blackmoore', 'Blackrock', 'Blackhand', 'Antonidas', 'Hyjal', 'Archimonde', 'Thrall', 'Eredar', 'Ysondre', 'Dalaran', 'Onyxia', 'Nefarian', 'The Maelstrom', 'Frostwolf', 'Aegwynn']
        realm_codes = ['KZ','TN','SM','DR','TM','RG','RV','AD','SV','FM','BB','BM','BR','BH','AN','HJ','AR','TH','ER','YS','DL','ON','NF','MA','FW','AE']
        row_offset = 32
        col_offset = 2

        for row in selection:
            for cell in row:
                c.execute("SELECT * FROM farmers WHERE realm=? AND account=?",
                        (realm_list[cell.row - row_offset], cell.col_idx - col_offset))
                db_row = c.fetchone()
                if db_row:
                    cell.value = realm_codes[cell.row - row_offset]
                    # char info
                    char_class = db_row[3]
                    intro_complete = True if db_row[4] == 1 else False
                    herbalism_rating = db_row[5] or 0
                    mining_rating = db_row[6] or 0
                    zin = db_row[7] or 0
                    os_deposit = db_row[8] or 0
                    os_seam = db_row[9] or 0

                    cell_bold = False if char_class == 'dh' else True
                    cell_underline = 'single' if char_class == 'dh' else 'none'
                    cell_color = "000000"

                    if zin == 3 and os_deposit == 3 and os_seam == 3:
                        cell_color = "14ccf5"
                    elif zin == 3 and os_deposit > 1 and os_seam > 1:
                        cell_color = "872bff"
                    elif zin == 3 and mining_rating >= 140:
                        cell_color = "1dde26"
                    elif zin == 3 and mining_rating < 140:
                        cell_color = "f54242"
                    elif zin == 2:
                        cell_color = "ff9412"
                    elif intro_complete == 1:
                        cell_color = "000000"
                    elif intro_complete == 0:
                        cell_color = "b8b8b8"

                    cell.font = Font(color=cell_color, bold=cell_bold, underline=cell_underline)

        wb.save(EXCEL_PATH)
        

if __name__ == '__main__':
    fdata = FarmerData()
    fdata.update_farmers()
    fdata.write_farmers_db()
    fdata.write_excel_table()

    print('success')