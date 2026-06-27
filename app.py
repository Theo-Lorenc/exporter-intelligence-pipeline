from flask import Flask, jsonify, send_file
import pandas as pd
import sqlite3
from config.settings import DB_FILE
from flask import Flask, jsonify, send_file
from flask import render_template

from run_pipeline import main  # reuse your pipeline


app = Flask(__name__)

@app.route("/suppliers")
def get_suppliers():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM exporters", conn)
    conn.close()

    return df.to_json(orient="records")


@app.route("/download-excel")
def download_excel():
    return send_file("data/exporters.xlsx", as_attachment=True)


@app.route("/refresh")
def refresh_data():
    main()  # run full pipeline
    return {"status": "updated"}


@app.route("/")
def home():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
