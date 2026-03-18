import os
import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def processar_sped(filepath):
    dados_0200 = []
    dados_h010 = []
    
    with open(filepath, 'r', encoding='latin1') as f:
        for linha in f:
            campos = linha.split('|')
            if len(campos) > 1:
                # Registro 0200: Cadastro de Itens
                if campos[1] == '0200':
                    dados_0200.append({
                        'COD_ITEM': campos[2],
                        'DESCR_ITEM': campos[3],
                        'UNID_INV': campos[6]
                    })
                # Registro H010: Inventário
                elif campos[1] == 'H010':
                    dados_h010.append({
                        'COD_ITEM': campos[2],
                        'UNID': campos[3],
                        'QTD': campos[4].replace(',', '.'),
                        'VL_UNIT': campos[5].replace(',', '.'),
                        'VL_ITEM': campos[6].replace(',', '.')
                    })

    df_0200 = pd.DataFrame(dados_0200)
    df_h010 = pd.DataFrame(dados_h010)
    
    # Cruza os dados para ter a descrição no inventário
    if not df_0200.empty and not df_h010.empty:
        df_final = pd.merge(df_h010, df_0200, on='COD_ITEM', how='left')
    else:
        df_final = df_h010

    output_path = os.path.join(UPLOAD_FOLDER, 'Inventario_SPED.xlsx')
    df_final.to_excel(output_path, index=False)
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "Nenhum arquivo enviado", 400
    
    file = request.files['file']
    if file.filename == '':
        return "Arquivo sem nome", 400

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        
        # Chama a função que cria o Excel
        excel_path = processar_sped(filepath)
        
        return send_file(excel_path, as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)