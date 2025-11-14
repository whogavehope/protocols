import sqlite3

def init_db():
    conn = sqlite3.connect("films.db")
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS films (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sidak_num TEXT UNIQUE NOT NULL,
        supplier_name TEXT NOT NULL,
        thickness REAL NOT NULL,
        thickness_score INTEGER GENERATED ALWAYS AS (
            CASE 
                WHEN thickness >= 0.35 THEN 5
                WHEN thickness BETWEEN 0.3 AND 0.34 THEN 4
                WHEN thickness BETWEEN 0.25 AND 0.29 THEN 3
                WHEN thickness BETWEEN 0.2 AND 0.24 THEN 2
                WHEN thickness < 0.19 THEN 1
                ELSE NULL
            END
        ) VIRTUAL,
        -- Новые дополнительные поля
        heating_score INTEGER CHECK(heating_score >= 1 AND heating_score <= 5),
        coffee_score INTEGER CHECK(coffee_score >= 1 AND coffee_score <= 5),
        oil_score INTEGER CHECK(oil_score >= 1 AND oil_score <= 5),
        -- ИЗМЕНЕНО: Твердость теперь цифровое значение '1'-'5' (хранится как текст)
        hardness TEXT CHECK(hardness IS NULL OR (hardness IN ('1', '2', '3', '4', '5'))), 
        
        -- ИЗМЕНЕНО: Логика hardness_score для цифровых значений
        hardness_score INTEGER GENERATED ALWAYS AS (
            CASE 
                WHEN hardness = '1' THEN 1
                WHEN hardness = '2' THEN 2
                WHEN hardness = '3' THEN 3
                WHEN hardness = '4' THEN 4
                WHEN hardness = '5' THEN 5
                ELSE NULL
            END
        ) VIRTUAL,
        
        ics_percent REAL GENERATED ALWAYS AS (
            CASE WHEN thickness_score IS NOT NULL 
                      AND heating_score IS NOT NULL 
                      AND coffee_score IS NOT NULL 
                      AND oil_score IS NOT NULL 
                      AND hardness_score IS NOT NULL
                     THEN ROUND((thickness_score + heating_score + coffee_score + oil_score + hardness_score) / 25.0 * 100, 2)
            END
        ) VIRTUAL,
        quality_level TEXT  -- Уровень качества
    )
    """)
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("✅ films.db инициализирована с дополнительными полями")