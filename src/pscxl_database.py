import logging
import sqlite3
import openpyxl

class PSCXL_Database:
    def __init__(self, db_name="pscxl.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.init_database()

    def init_database(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS workbooks (
                workbook_name TEXT, 
                sheet_name TEXT, 
                value TEXT, 
                quantity INTEGER,
                stacked INTEGER,
                PRIMARY KEY (workbook_name, sheet_name, value)
            )
        ''')
        self.conn.commit()

    def insert_data(self, workbook_name, sheet_name, value, quantity, stacked):
        try:
            self.cursor.execute('''
                INSERT OR IGNORE INTO workbooks (workbook_name, sheet_name, value, quantity, stacked)
                VALUES (?, ?, ?, ?, ?)
            ''', (workbook_name, sheet_name, value, quantity, stacked))
            self.conn.commit()
        except sqlite3.IntegrityError as e:
            logging.error(f"Error inserting data: {e}")

    def clear_workbook_data(self, workbook_name):
        try:
            self.cursor.execute('''
                DELETE FROM workbooks WHERE workbook_name = ?
            ''', (workbook_name,))
            self.conn.commit()
        except sqlite3.Error as e:
            logging.error(f"Error clearing data: {e}")

    def update_data(self, workbook_name, sheet_name, value, new_value, new_quantity, new_stacked):
        self.cursor.execute('''
            UPDATE workbooks
            SET value = ?, quantity = ?, stacked = ?
            WHERE workbook_name = ? AND sheet_name = ? AND value = ?
        ''', (new_value, new_quantity, new_stacked, workbook_name, sheet_name, value))
        self.conn.commit()

    def delete_data(self, workbook_name, sheet_name, value):
        self.cursor.execute('''
            DELETE FROM workbooks
            WHERE workbook_name = ? AND sheet_name = ? AND value = ?
        ''', (workbook_name, sheet_name, value))
        self.conn.commit()

    def fetch_workbooks(self):
        self.cursor.execute('SELECT DISTINCT workbook_name FROM workbooks')
        return [row[0] for row in self.cursor.fetchall()]

    def fetch_sheets(self, workbook_name):
        self.cursor.execute('SELECT DISTINCT sheet_name FROM workbooks WHERE workbook_name = ?', (workbook_name,))
        return [row[0] for row in self.cursor.fetchall()]

    def fetch_data(self, workbook_name, sheet_name):
        self.cursor.execute('SELECT value, quantity, stacked FROM workbooks WHERE workbook_name = ? AND sheet_name = ?', 
                            (workbook_name, sheet_name))
        return self.cursor.fetchall()
