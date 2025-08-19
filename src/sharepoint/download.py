import os, requests

def list_folder_files(drive_id: str, folder_path: str, headers: dict[str, str]) -> list:
    """Lista os arquivos em uma pasta específica no SharePoint."""
    
    folder_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"

    response = requests.get(folder_url, headers=headers)

    if response.status_code != 200:
        raise Exception(f"Erro ao listar arquivos na pasta '{folder_path}': {response.json()}")

    return response.json().get('value', [])

def download_file_from_sharepoint(access_token: str, drive_id: str, sharepoint_path: str, local_file_path: str, files=[]) -> None:
    """Faz download de um arquivo de uma biblioteca específica no SharePoint."""
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json',
    }
    
    items = list_folder_files(drive_id, os.path.dirname(sharepoint_path), headers)
    if not items:
        raise Exception(f"Nenhum arquivo encontrado na pasta '{os.path.dirname(sharepoint_path)}'.")
    
    # Filtra apenas os arquivos cujo nome (antes do "_") está em files
    filtered_items = [
        item for item in items
        if item.get('file') and item.get('name', '').split('_')[0].replace(' ', '').upper() in [f for f in files]
    ]

    for item in filtered_items:
        download_url = item.get('@microsoft.graph.downloadUrl')
        file = item.get('name')
        resp = requests.get(download_url, headers=headers)
        resp.raise_for_status()
        with open(os.path.join(local_file_path, file), 'wb') as f:
            f.write(resp.content)