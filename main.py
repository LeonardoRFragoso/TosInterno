import subprocess
import sys
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

def upload_files():
    # Escopos de acesso à API do Drive
    SCOPES = ['https://www.googleapis.com/auth/drive']
    SERVICE_ACCOUNT_FILE = r"C:\Users\leonardo.fragoso\Desktop\Projetos\Depot-Project\gdrive_credentials.json"
    
    # Cria as credenciais com o arquivo de serviço
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    
    drive_service = build('drive', 'v3', credentials=credentials)
    
    # ID da pasta no Google Drive onde os arquivos serão enviados
    folder_id = "1ROPmQRq9Wy_Ugzi9rZQ2Vnt5mfnSZ0a5"
    
    downloads_folder = os.path.join(os.getcwd(), "downloads")

    # Procura arquivos Excel na pasta "downloads"
    for filename in os.listdir(downloads_folder):
        if filename.endswith(".xlsx"):
            filepath = os.path.join(downloads_folder, filename)

            # Verifica se o arquivo já existe no Drive (mesmo nome e mesma pasta)
            query = (
                f"name = '{filename}' "
                f"and '{folder_id}' in parents "
                f"and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
                f"and trashed = false"
            )
            response = drive_service.files().list(q=query, fields="files(id, name)").execute()
            files_found = response.get('files', [])

            # Prepara o conteúdo do arquivo para envio
            media = MediaFileUpload(
                filepath,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                resumable=True
            )

            if files_found:
                # Se o arquivo já existe, faz o update mantendo o mesmo ID/URL
                file_id = files_found[0]['id']
                updated_file = drive_service.files().update(
                    fileId=file_id,
                    media_body=media
                ).execute()
                print(f"Arquivo '{filename}' atualizado no Drive. ID: {updated_file.get('id')}")
            else:
                # Caso não exista, cria um novo arquivo no Drive
                file_metadata = {
                    'name': filename,
                    'parents': [folder_id]
                }
                new_file = drive_service.files().create(
                    body=file_metadata,
                    media_body=media,
                    fields='id'
                ).execute()
                print(f"Arquivo '{filename}' criado no Drive. ID: {new_file.get('id')}")

def main():
    # Executa os scripts na ordem desejada
    subprocess.run([sys.executable, "rbt.py"], check=True)
    #subprocess.run([sys.executable, "multirio.py"], check=True)
    #subprocess.run([sys.executable, "tecon.py"], check=True)
    
    # Envia as planilhas da pasta "downloads" para a pasta específica do Google Drive
    upload_files()

if __name__ == "__main__":
    main()
