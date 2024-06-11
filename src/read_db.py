import sqlite3

def read_database(db_path):
    # Connect to the database
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Get the list of tables
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = cursor.fetchall()

    for table in tables:
        table_name = table[0]
        print(f"Table: {table_name}")

        # Get the table's columns
        cursor.execute(f"PRAGMA table_info({table_name});")
        columns = [description[1] for description in cursor.fetchall()]
        print("Columns:", ", ".join(columns))

        # Get the table's data
        cursor.execute(f"SELECT * FROM {table_name};")
        rows = cursor.fetchall()
        for row in rows:
            print(row)
        
        print("-" * 40)

    # Close the connection
    conn.close()

if __name__ == "__main__":
    db_path = 'pscxl.db'  # Path to your SQLite database file
    read_database(db_path)
