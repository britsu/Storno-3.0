import os
from flask import Flask, request, render_template, jsonify, send_from_directory
from werkzeug.utils import secure_filename
from .engine import processar_planilha

app = Flask(__name__)

# Configurações de upload
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__name__)), 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024 # Max 16MB

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'csv', 'xlsx'}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    if 'file' not in request.files:
        return jsonify({"sucesso": False, "erro": "Nenhum arquivo enviado."})
        
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({"sucesso": False, "erro": "Nenhum arquivo selecionado."})
        
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Chama a engine para processar
        resultado = processar_planilha(filepath, app.config['OUTPUT_FOLDER'])
        
        # Remove arquivo original após processar
        try:
            os.remove(filepath)
        except:
            pass
            
        return jsonify(resultado)
    else:
        return jsonify({"sucesso": False, "erro": "Formato de arquivo não suportado. Use .xlsx ou .csv"})

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    # INSTRUÇÕES DE EXECUÇÃO LOCAL:
    # 1. pip install -r requirements.txt
    # 2. python app.py
    # 3. Abrir http://localhost:5000 no navegador
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
