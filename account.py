import sqlite3


conn = sqlite3.connect('instance/db.sqlite')
cursor = conn.cursor()


conn.close()