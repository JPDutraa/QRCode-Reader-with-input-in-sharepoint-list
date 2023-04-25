# Projeto QRCode Scan - Sharepoint

Esse projeto consiste em uma aplicação Pytho n que utiliza streaming de câmera para capturar QrCodes e registrar automaticamente em uma lista online do Sharepoint.

## Funcionalidades
- Captura de QrCode via streaming de câmera
- Registro automático dos dados do QRCode em uma lista online do Sharepoint
- Autentificação no Sharepoint via OAuth 2.0

### Requisitos
- Python 3.7 ou superior
- Pip (Gerenciador de pacotes do Python)


### Dependências
- opencv-python
- pyzbar
- office365-connector
- winsound
- logging

## Instalação

1. Clone esse repositório:
```
git clone https://github.com/JPDutraa/QRCode-Reader-with-input-in-sharepoint-list.git
```
2. Acesse a pasta do projeto:
```
cd QRCode-Reader-with-input-in-sharepoint-list
```
3. Instale as dependências:
```
pip install -r requeriments.txt
```

## Uso

Execute o script principal
```
python LeitorQRCode-Sharepoint.py
```

O aplicativo iniciará o streaming da câmera e começará a procurar por QRCodes. Quando um QRCode for detectado, os dados serão registrados automaticamente na lista especificada do SharePoint.

## Funcionamento do script

1. Importa as bibliotecas necessárias.
2. Configura o logger para gravar logs no arquivo "log.txt".
3. Define as credenciais de acesso ao SharePoint e cria o contexto de conexão.
4. Define o tempo mínimo de espera entre a leitura de dois QR codes consecutivos.
5. Captura a transmissão ao vivo de uma câmera.
6. Verifica se a transmissão foi capturada com sucesso e encerra o script em caso de erro.
7. Inicia um loop para capturar e processar frames da transmissão.
8. Converte a imagem capturada para tons de cinza.
9. Busca por QR codes na imagem.
10. Se encontrar um QR code, verifica se o tempo mínimo de espera foi atendido.
11. Adiciona um item na lista do SharePoint com os dados do QR code.
12. Toca um som e exibe uma mensagem de confirmação na tela.
13. Exibe o frame capturado na janela de exibição.
14. Encerra o loop e libera os recursos quando a tecla 'q' for pressionada.