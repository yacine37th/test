from flask import Flask, request
import pandas as pd, os

app = Flask(__name__)
EXCEL_FILE = "donnees.xlsx"

def save_to_excel(new_data):
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df = pd.concat([df, pd.DataFrame([new_data])], ignore_index=True)
    else:
        df = pd.DataFrame([new_data])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/receive', methods=['POST'])
def receive_data():
    data = request.json
    if not data:
        return {"status": "error", "message": "No JSON received"}, 400
    save_to_excel(data)
    return {"status": "success", "message": "Data saved to Excel"}

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
