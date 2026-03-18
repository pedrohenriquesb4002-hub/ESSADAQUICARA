from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400
    file = request.files['file']
    path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(path)
    itens, produtos = [], {}
    try:
        with open(path, "r", encoding="latin1") as f:
            for linha in f:
                p = linha.strip().split("|")
                if len(p) < 2: continue
                if p[1] == "0200": produtos[p[2]] = p[3]
                if p[1] == "H010":
                    itens.append({
                        "COD_ITEM": p[2], "DESCRICAO": produtos.get(p[2], ""),
                        "UNID": p[3], "QTD": float(p[4].replace(",", ".")),
                        "VL_UNIT": float(p[5].replace(",", ".")), "VL_ITEM": float(p[6].replace(",", "."))
                    })
        df = pd.DataFrame(itens)
        excel_path = os.path.join(UPLOAD_FOLDER, "inventario_final.xlsx")
        df.to_excel(excel_path, index=False)
        return send_file(excel_path, as_attachment=True)
    finally:
        if os.path.exists(path): os.remove(path)

import os

if __name__ == '__main__':
    # O Render/Railway usa uma porta variável, por isso usamos o 'PORT'
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)