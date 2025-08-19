from .auth import authenticate
from .site import get_site_id, get_drive_id
from .upload import upload_file_to_sharepoint
from .download import download_file_from_sharepoint

from dataclasses import dataclass

__all__ = [
    "Sharepoint",
]

class Sharepoint:
    def __init__(self, settings):
        self.config = SharepointConfig(
            SHAREPOINT_SITE_URL=settings.SHAREPOINT_SITE_URL,
            CREDENCIAL=Credentials(
                CLIENT_ID=settings.CLIENT_ID,
                CLIENT_SECRET=settings.CLIENT_SECRET,
                TENANT_ID=settings.TENANT_ID,
                USERNAME=settings.USERNAME,
                PASSWORD=settings.PASSWORD,
                SCOPES=settings.SCOPES.split(),
            )
        )
        self.AUTHORITY = self.config.CREDENCIAL.AUTHORITY

        self.access_token = authenticate(
            self.config.CREDENCIAL.CLIENT_ID,
            self.AUTHORITY,
            self.config.CREDENCIAL.CLIENT_SECRET,
            self.config.CREDENCIAL.USERNAME,
            self.config.CREDENCIAL.PASSWORD,
            self.config.CREDENCIAL.SCOPES)

    def upload(self, local_file_path,sharepoint_path, biblioteca="Documentos"):
        """Faz upload de um arquivo para o SharePoint."""
        try:
            site_id = get_site_id(self.access_token, self.config.SHAREPOINT_SITE_URL)
            drive_id = get_drive_id(self.access_token, site_id, biblioteca)

            upload_file_to_sharepoint(self.access_token, drive_id, local_file_path, sharepoint_path)

        except Exception as e:
           pass
    
    def download(self, sharepoint_path, local_file_path="Temp", biblioteca="Documentos", files = []):
        """Faz download de um arquivo do SharePoint."""
        
        try:
            site_id = get_site_id(self.access_token, self.config.SHAREPOINT_SITE_URL)
            drive_id = get_drive_id(self.access_token, site_id, biblioteca)

            download_file_from_sharepoint(self.access_token, drive_id, sharepoint_path, local_file_path,files=files)
            
        except Exception as e:
            pass

@dataclass
class Credentials:
    CLIENT_ID: str
    CLIENT_SECRET: str
    TENANT_ID: str
    USERNAME: str
    PASSWORD: str
    SCOPES: list

    @property
    def AUTHORITY(self) -> str:
        return f"https://login.microsoftonline.com/{self.TENANT_ID}"

@dataclass
class SharepointConfig:
    SHAREPOINT_SITE_URL: str
    CREDENCIAL: Credentials