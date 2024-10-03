from flask import Flask, render_template_string
import pandas as pd

app = Flask(__name__)

# Route untuk menampilkan data Excel di localhost
@app.route('/')
def display_excel():
    # Membaca file Excel
    filename = "data_belanja.xlsx"
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        return "File Excel tidak ditemukan."

    # Mengonversi DataFrame ke HTML
    table_html = df.to_html(index=False, classes='table table-bordered', border=0)

    # Template HTML sederhana untuk menampilkan data
    template = '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Data Belanja</title>
        <style>
            body { font-family: Arial, sans-serif; margin: 40px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { padding: 8px 12px; text-align: center; border: 1px solid #ddd; }
            th { background-color: #f2f2f2; }
            tr:hover { background-color: #f5f5f5; }
        </style>
    </head>
    <body>
        <h2>Database Penjualan Barang</h2>
        {{ table_html | safe }}
    </body>
    </html>
    '''
    # Render HTML dengan tabel dari DataFrame
    return render_template_string(template, table_html=table_html)

if __name__ == '__main__':
    app.run(debug=True)

