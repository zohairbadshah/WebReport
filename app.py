from flask import Flask, render_template, request, redirect, url_for,send_file
import requests
import pandas as pd

from openpyxl import Workbook
from io import BytesIO
import json
app = Flask(__name__)
final_data=None
final_report_type=None
# Function to make API requests and return data as a DataFrame
def make_api_request(api_url):
    response = requests.post(api_url, files={"file": ("file.xlsx", request.files["file"].read())})
    if response.status_code == 200:
        data = response.json().get("data", [])
        df = pd.DataFrame(data)
        return df
    else:
        return None


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    global final_data
    global final_report_type
    file = request.files["file"]
    report_type = request.form["report_type"]

    if file and report_type:
        if report_type == "overall":
            final_report_type="overall"
            api_url = "http://164.52.205.152:8085/over_all_report_button"
            df = make_api_request(api_url)

            if df is not None:
                final_data=df
                return render_template("overall.html", data=df.to_dict(orient="records"))
            else:
                return "Error: Unable to fetch data from the API."

        elif report_type == "daily":
            final_report_type="daily"
            api_url = "http://164.52.205.152:8085/daily_batch_report_button"
            df = make_api_request(api_url)
            final_data=df
            if df is not None:
                return render_template("daily.html", data=df.to_dict(orient="records"))
            else:
                return "Error: Unable to fetch data from the API."

    return redirect(url_for("index"))

@app.route("/download_excel", methods=["POST"])
def download_excel():
    if final_data is not None:
        # Create an Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_data.to_excel(writer, sheet_name='Sheet1', index=False)

        output.seek(0)
        return send_file(output, as_attachment=True,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', download_name=f'{final_report_type}.xlsx')
    else:
        print("Error")

if __name__ == "__main__":
    app.run(debug=False,host='0.0.0.0')
