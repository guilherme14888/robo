from flask import Flask, render_template, request, send_file, after_this_request
import os
from concurrent.futures import ThreadPoolExecutor
import boleto

app = Flask(__name__)
executor = ThreadPoolExecutor(max_workers=2)

def executar_robo(contrato: str, email: str, tipo: str) -> str:
    """Executa o robô principal e retorna o caminho do PDF gerado."""
    return boleto.main(contrato, email)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/cnvw', methods=['GET', 'POST'])
def cnvw():
    if request.method == 'POST':
        contrato = request.form.get('contratos') or request.form.get('contrato')
        email = request.form.get('email')
        
        if not contrato or not email:
            return "⚠️ Campo contrato ou email não informado.", 400

        future = executor.submit(executar_robo, contrato, email, 'CNVW')
        pdf_path = future.result()

        @after_this_request
        def remove_file(response):
            try:
                os.remove(pdf_path)
            except OSError:
                pass
            return response

        return send_file(pdf_path, download_name=f"{contrato}.pdf", as_attachment=True)
    
    return render_template('form.html', titulo='CNVW')

@app.route('/fiat', methods=['GET', 'POST'])
def fiat():
    if request.method == 'POST':
        contrato = request.form.get('contratos') or request.form.get('contrato')
        email = request.form.get('email')

        if not contrato or not email:
            return "⚠️ Campo contrato ou email não informado.", 400

        future = executor.submit(executar_robo, contrato, email, 'FIAT')
        pdf_path = future.result()

        @after_this_request
        def remove_file(response):
            try:
                os.remove(pdf_path)
            except OSError:
                pass
            return response

        return send_file(pdf_path, download_name=f"{contrato}.pdf", as_attachment=True)
    
    return render_template('form.html', titulo='FIAT')

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
