import json
import os

dirname = os.path.dirname(__file__)
SETTINGS_FILE = os.path.join(dirname, 'settings.json')

def set_setting(field):
    return settings['settings'].get(field, None) or settings['default_settings'].get(field, None)
    
with open(SETTINGS_FILE) as f:
    settings = json.load(f)


# characters database
CHAR_DB = set_setting('char_db')
REALMS = set_setting('realms')
AUCTIONS = set_setting('auctions')
LUA_PATHS = set_setting('lua_paths')
LUA_FILES = set_setting('lua_files')
FARMERS_DB = set_setting('farmers_db')
EXCEL_PATH = set_setting('excel_path')