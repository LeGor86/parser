import sqlite3, os

db_path = r'C:\Users\LeGorUSER\Desktop\WebDev\Parser_for_site\!ParserV4-7\user_status.db'

conn = sqlite3.connect(db_path)
cur = conn.cursor()
cur.execute('''
CREATE TABLE IF NOT EXISTS user_status (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    url TEXT NOT NULL,
    status TEXT NOT NULL,
    timestamp DATETIME NOT NULL,
    online_indicators TEXT,
    success BOOLEAN,
    error TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)
''')
cur.execute('CREATE INDEX IF NOT EXISTS idx_url_timestamp ON user_status (url, timestamp)')
conn.commit()
conn.close()

print('Database successfully created')
