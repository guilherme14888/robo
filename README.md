# Robo Contratos

Interface web simples para iniciar o robô de automação e receber o resultado em PDF.
Utiliza [Bootstrap](https://getbootstrap.com/) para uma aparência agradável.

## Requisitos

- Python 3.9+
- [Flask](https://flask.palletsprojects.com/)
- Dependências do script `robocontratos.py` (Selenium, etc.)
- Arquivo `smtp_config.py` configurado com servidor, porta e credenciais de e-mail
- O robô opera em segundo plano utilizando uma janela virtual **1920x1080** sem alterar a resolução do seu monitor

Instale as dependências com:

```bash
pip install flask
```

## Executando

```bash
python app.py
```

O servidor ficará disponível para a rede em `http://<seu_ip>:5000`.
Escolha **CNVW** ou **FIAT**, informe o número do contrato e seu e-mail e clique em **Enviar**.
O robô será executado em segundo plano e, ao finalizar, o PDF ficará disponível
para download temporário e enviado por e-mail para o endereço informado.
Configure suas credenciais de envio no arquivo `smtp_config.py` antes de iniciar a aplicação.