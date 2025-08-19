import os, requests

def upload_file_to_sharepoint(access_token: str, drive_id: str, local_file_path: str, sharepoint_path: str) -> None:
    """Faz upload de um arquivo para uma biblioteca específica no SharePoint."""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
    }
    file_name = os.path.basename(local_file_path)
    file_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sharepoint_path}/{file_name}:/content"

    with open(local_file_path, 'rb') as file_data:
        response = requests.put(file_url, headers=headers, data=file_data)

    if response.status_code != 201:
        raise Exception(f"Erro ao fazer upload do arquivo '{file_name}': {response.json()}")
    else:
        print(f"Arquivo '{file_name}' enviado com sucesso.")