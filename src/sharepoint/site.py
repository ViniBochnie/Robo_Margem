import requests

def get_site_id(access_token: str, site_info_url: str) -> str:
    """Obtém o site ID do SharePoint usando a URL do site."""
    headers = {'Authorization': f'Bearer {access_token}'}
    site_info_url_format = f"https://graph.microsoft.com/v1.0/sites/{site_info_url}"
    site_info_response = requests.get(site_info_url_format, headers=headers)

    if site_info_response.status_code == 200:
        return site_info_response.json()["id"]
    else:
        raise Exception("Erro ao obter o site ID:", site_info_response.json())
    
def get_drive_id(access_token: str, site_id: str, library_name: str) -> str:
    """Obtém o drive_id de uma biblioteca de documentos específica no SharePoint usando $select."""
    headers = {'Authorization': f'Bearer {access_token}'}
    drive_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives?$select=id,name"
    drive_response = requests.get(drive_url, headers=headers)

    if drive_response.status_code == 200:
        drives = drive_response.json().get("value", [])
        for drive in drives:
            if drive["name"] == library_name:
                return drive["id"]
        raise Exception(f"Biblioteca '{library_name}' não encontrada.")
    else:
        raise Exception("Erro ao obter o drive ID:", drive_response.json())