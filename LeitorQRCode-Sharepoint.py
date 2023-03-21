import cv2
from pyzbar import pyzbar
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.lists.list import ListItem
import winsound
import time


# Define as credenciais de acesso ao SharePoint
username = 'Email'
password = 'Senha'
site_url = 'Site do Sharepoint'
list_name = 'Nome da Lista do Sharepoint'

# Cria o contexto de conexão com o SharePoint
ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

# Define o tempo de espera mínimo entre a leitura de dois QR codes consecutivos
min_time_interval = 5  # segundos

# Captura a transmissão ao vivo
stream_url = 'URL do Video'
cap = cv2.VideoCapture(stream_url)

# Verifica se a transmissão foi capturada com sucesso
if not cap.isOpened():
    print("Erro ao capturar transmissão ao vivo")
    exit()

# Loop para capturar e processar frames da transmissão
last_barcode_time = 0  # Inicializa a última vez que um QR code foi lido

# Loop para capturar e processar frames da transmissão
while True:
    ret, frame = cap.read()
    if not ret:
        print("Erro ao capturar frame da transmissão")
        break

    # Converte a imagem capturada para tons de cinza
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Busca por QRCode na imagem
    barcodes = pyzbar.decode(gray)

    # Se encontrar um QRCode, adiciona um item na lista do SharePoint
    if len(barcodes) > 0:
        barcode_data = barcodes[0].data.decode('utf-8')
        current_time = time.time() #Obtém o tempo real
        if current_time - last_barcode_time >= min_time_interval:
             # Tempo suficiente passou desde a última leitura de QR code
            last_barcode_time = current_time # Atualiza o tempo da última leitura
            print(f"QRCode lido: {barcode_data}")
            barcode_data_split = barcode_data.split(";")
            barcode_data_strip_0 = barcode_data_split[0].strip()
            barcode_data_strip_1 = barcode_data_split[1].strip()
            barcode_data_strip_2 = barcode_data_split[2].strip()
            barcode_data_strip_3 = barcode_data_split[3].strip()
            print(barcode_data_strip_0)
            print(barcode_data_strip_1)
            print(barcode_data_strip_2)
            print(barcode_data_strip_3)

            list_item = ctx.web.lists.get_by_title(list_name).add_item({
                'Title': barcode_data_strip_0,
                'NOME': barcode_data_strip_1,
                'CIDADE': barcode_data_strip_2,
                'PROCESSO': barcode_data_strip_3
            })
            ctx.execute_query()
            print(f"Item adicionado na lista {list_name} com sucesso!")

            # Toca um som para indicar que um QR code foi lido
            winsound.PlaySound("SystemQuestion", winsound.SND_ALIAS)

            #  # Exibe uma mensagem de confirmação na tela
            message = f"QR Code lido:\n{barcode_data_strip_0}\n{barcode_data_strip_1}\n{barcode_data_strip_2}\n{barcode_data_strip_3}"
            cv2.putText(frame, message, (50, 50), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
            # cv2.imshow("QRCode Lido com Sucesso", frame)
            cv2.waitKey(3000)  # Exibe a mensagem por 3 segundos
    
    # Exibe o frame capturado na janela de exibição
    cv2.imshow("LIVESTREAM - WEBCAM", frame)
    
    # Espera por uma tecla e sai se a tecla 'q' for pressionada
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
