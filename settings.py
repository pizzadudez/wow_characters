import json

def set_setting(field):
    return settings['settings'].get(field, None) or settings['default_settings'].get(field, None)
    
with open('settings.json') as f:
    settings = json.load(f)


# characters database
CHAR_DB = set_setting('char_db')
LUA_PATHS = set_setting('lua_paths')
FARMERS_DB = set_setting('farmers_db')
EXCEL_PATH = set_setting('excel_path')