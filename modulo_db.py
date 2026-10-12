import sqlite3

def conectar():
    return sqlite3.connect("informes.db")

def crear_tablas():
    con = conectar()
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS informes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            tipo TEXT,
            obra TEXT,
            fecha TEXT,
            docx TEXT,
            pdf TEXT
        )
    """)
    con.commit()
    con.close()

def guardar_informe(codigo, tipo, obra, fecha, docx, pdf):
    con = conectar()
    cur = con.cursor()
    cur.execute("""
        INSERT INTO informes (codigo, tipo, obra, fecha, docx, pdf)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (codigo, tipo, obra, fecha, docx, pdf))
    con.commit()
    con.close()
