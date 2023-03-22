# Projeto QRCode Scan - Sharepoint
[![N|Solid](https://cldup.com/dTxpPi9lDf.thumb.png)](https://nodesource.com/products/nsolid)
[![Build Status](https://travis-ci.org/joemccann/dillinger.svg?branch=master)](https://travis-ci.org/joemccann/dillinger)

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

