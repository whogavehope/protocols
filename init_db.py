import sqlite3

def init_db():
    conn = sqlite3.connect("films.db")
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS films (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sidak_num TEXT UNIQUE NOT NULL,
        supplier_name TEXT NOT NULL,
        thickness REAL NOT NULL
    )
    """)
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("✅ films.db инициализирована")
