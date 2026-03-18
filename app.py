import os
import pandas as pd
from flask import Flask, render_template, request, send_file

app = Flask(__name__)

# Configuração de pastas
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def processar_sped(filepath):
    dados_0200 = []
    dados_h010 = []
    
    # Lendo o arquivo SPED
    with open(filepath, 'r', encoding='latin1') as f:
        for linha in f:
            campos = linha.split('|')
            if len(campos) > 1:
                # Registro 0200: Cadastro de Itens (Pega Código e Descrição)
                if campos[1] == '0200':
                    dados_0200.append({
                        'COD_ITEM': campos[2],
                        'DESCRICAO': campos[3]
                    })
                # Registro H010: Inventário (Pega os valores)
                elif campos[1] == 'H010':
                    dados_h010.append({
                        'COD_ITEM': campos[2],
                        'UNID': campos[3],
                        'QTD': campos[4].replace(',', '.'),
                        'VL_UNIT': campos[5].replace(',', '.'),
                        'VL_ITEM': campos[6].replace(',', '.')
                    })

    # Transformando em Tabelas (DataFrames)
    df_0200 = pd.DataFrame(dados_0200)
    df_h010 = pd.DataFrame(dados_h010)
    
    if not df_0200.empty and not df_h010.empty:
        # Cruza os dados: coloca a Descrição do 0200 ao lado do Código do H010
        df_final = pd.merge(df_h010, df_0200, on='COD_ITEM', how='left')
        
        # Define a ordem EXATA das colunas igual à sua imagem
        colunas_ordenadas = ['COD_ITEM', 'DESCRICAO', 'UNID', 'QTD', 'VL_UNIT', 'VL_ITEM']
        
        # Garante que só existam essas colunas e nessa ordem
        df_final = df_final[colunas_ordenadas]
        
        # Converte os textos para números reais para o Excel entender como valor
        df_final['QTD'] = pd.to_numeric(df_final['QTD'], errors='coerce')
        df_final['VL_UNIT'] = pd.to_numeric(df_final['VL_UNIT'], errors='coerce')
        df_final['VL_ITEM'] = pd.to_numeric(df_final['VL_ITEM'], errors='coerce')
    else:
        df_final = df_h010

    # Nome do arquivo de saída
    output_filename = 'Inventario_SPED_Formatado.xlsx'
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)
    
    # Salva o Excel (index=False remove aquela coluna de números 0, 1, 2 na esquerda)
    df_final.to_excel(output_path, index=False)
    return output_path

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "Erro: Nenhum arquivo enviado", 400
    
    file = request.files['file']
    if file.filename == '':
        return "Erro: Arquivo sem nome", 400

    if file:
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        
        # Executa a lógica de conversão
        try:
            excel_path = processar_sped(filepath)
            return send_file(excel_path, as_attachment=True)
        except Exception as e:
            return f"Erro ao processar o arquivo: {str(e)}", 500

if __name__ == '__main__':
    # Configuração necessária para rodar tanto no seu PC quanto no Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)