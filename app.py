from flask import Flask, render_template, request, redirect, url_for, send_file
import sqlite3
import pandas as pd
import io
from datetime import datetime

app = Flask(__name__)

# Banco de Dados
def init_sqlite_db():
    conn = sqlite3.connect('database.db')
    print("Banco de dados aberto com sucesso")
    conn.execute('''
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            address TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL
        );
    ''')
    print("Tabela criada com sucesso")
    conn.close()

init_sqlite_db()

#ROTAS

@app.route('/')
def index():
    return render_template('home.html')

@app.route('/add-client/', methods=['GET', 'POST'])
def add_client():
    if request.method == 'POST':
        try:
            name = request.form['name']
            address = request.form['address']
            phone = request.form['phone']
            email = request.form['email']

            with sqlite3.connect('database.db') as con:
                cur = con.cursor()
                cur.execute("INSERT INTO clients (name, address, phone, email) VALUES (?, ?, ?, ?)",
                            (name, address, phone, email))
                con.commit()
                msg = "Cliente adicionado com sucesso!"
        except:
            con.rollback()
            msg = "Erro ao adicionar o cliente"
        finally:
            return render_template('result.html', msg=msg)
            con.close()
    return render_template('add_client.html')

@app.route('/search-client/', methods=['GET', 'POST'])
def search_client():
    if request.method == 'POST':
        name = request.form['name']
        con = sqlite3.connect('database.db')
        con.row_factory = sqlite3.Row
        cur = con.cursor()
        cur.execute("SELECT * FROM clients WHERE name LIKE ?", ('%'+name+'%',))
        rows = cur.fetchall()
        return render_template('view_clients.html', rows=rows)
    return render_template('search_client.html')

@app.route('/export-excel/', methods=['POST'])
def export_excel():
    name = request.form.get('name')
    con = sqlite3.connect('database.db')
    df = pd.read_sql_query("SELECT * FROM clients WHERE name LIKE ?", con, params=('%'+name+'%',))
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Clientes')
    writer.close()
    output.seek(0)
    con.close()
    return send_file(output, download_name='clientes.xlsx', as_attachment=True)


# Ano atual
@app.context_processor
def inject_now():
    return {'current_year': datetime.utcnow().year}

if __name__ == '__main__':
    app.run(debug=True)
