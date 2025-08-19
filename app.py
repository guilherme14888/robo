from flask import Flask, render_template, request, send_file, after_this_request
import os
from concurrent.futures import ThreadPoolExecutor
import boleto
import contrato as contrato_script

app = Flask(__name__)
executor = ThreadPoolExecutor(max_workers=2)

def run_boleto(contrato_num: str) -> str:
    """Executa o script de boletos."""
    return boleto.main(contrato_num)

def run_contrato(contrato_num: str) -> str:
    """Executa o script de contratos."""
    return contrato_script.main(contrato_num)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/boleto', methods=['GET', 'POST'])
def boleto_view():
    if request.method == 'POST':
        contrato_num = request.form.get('contrato')
        if not contrato_num:
            return "⚠️ Campo contrato não informado.", 400
        future = executor.submit(run_boleto, contrato_num)
        pdf_path = future.result()
        @after_this_request
        def remove_file(response):
            try:
                os.remove(pdf_path)
            except OSError:
                pass
            return response
        return send_file(pdf_path, download_name=f"{contrato_num}.pdf", as_attachment=True)
    return render_template('form.html', titulo='Boleto')

@app.route('/contrato', methods=['GET', 'POST'])
def contrato_view():
    if request.method == 'POST':
        contrato_num = request.form.get('contrato')
        if not contrato_num:
            return "⚠️ Campo contrato não informado.", 400
        future = executor.submit(run_contrato, contrato_num)
        pdf_path = future.result()
        @after_this_request
        def remove_file(response):
            try:
                os.remove(pdf_path)
            except OSError:
                pass
            return response
        return send_file(pdf_path, download_name=f"{contrato_num}.pdf", as_attachment=True)
    return render_template('form.html', titulo='Contrato')

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)
