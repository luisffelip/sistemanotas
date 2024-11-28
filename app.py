from flask import Flask, request, send_file, render_template
import pandas as pd
import os

app = Flask(__name__)
app = Flask(__name__, static_folder='static')

# Função para corrigir as provas (sua lógica atual)
def corrigir_provas(gabarito_path, respostas_path):
    # Carregar os dados do gabarito e das respostas dos alunos
    gabarito_df = pd.read_excel(gabarito_path)
    respostas_alunos_df = pd.read_excel(respostas_path)

    # Converter o gabarito para um dicionário
    gabarito_dict = gabarito_df.set_index('Questão')['Resposta'].to_dict()

    # Corrigir as provas e calcular as notas
    notas = []
    for idx, row in respostas_alunos_df.iterrows():
        nome = row['ID']
        respostas = row[1:].to_dict()  # Pular a coluna do nome
        nota = 0
        for questao, resposta in respostas.items():
            if gabarito_dict.get(questao) == resposta:
                nota += 1
        notas.append({'ID': row['ID'], 'Nota': nota})

    # Criar um DataFrame com as notas
    notas_df = pd.DataFrame(notas)
    output_path = "notas_alunos.xlsx"
    notas_df.to_excel(output_path, index=False, sheet_name='Notas')

    return output_path

# Rota principal (carrega a página)
@app.route("/")
def index():
    return render_template("index.html")

#Rota para processar as imagens 
@app.route('/')
def home():
    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)

# Rota para receber os arquivos e processá-los
@app.route("/upload", methods=["POST"])
def upload_files():
    gabarito_file = request.files['gabarito']
    respostas_file = request.files['respostas']

    # Salvar os arquivos enviados
    gabarito_path = "gabarito_temp.xlsx"
    respostas_path = "respostas_temp.xlsx"
    gabarito_file.save(gabarito_path)
    respostas_file.save(respostas_path)

    try:
        # Corrigir as provas e gerar o arquivo final
        resultado_path = corrigir_provas(gabarito_path, respostas_path)
        return send_file(resultado_path, as_attachment=True)
    except Exception as e:
        return f"Ocorreu um erro: {str(e)}", 500
    finally:
        # Limpar arquivos temporários
        os.remove(gabarito_path)
        os.remove(respostas_path)

if __name__ == "__main__":
    app.run(debug=True)
