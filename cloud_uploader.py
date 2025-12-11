import os
import json
import requests
import msal
from abc import ABC, abstractmethod

# Google deps (import inside class or try/except to avoid crashing if missing)
try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
except ImportError:
    pass

class CloudUploader(ABC):
    @abstractmethod
    def upload_file(self, file_path):
        """Uploads file and returns a shareable link."""
        pass

class GDriveUploader(CloudUploader):
    def __init__(self, credentials_path, root_folder_name="AutoArchive", sub_folder_name="Scanned"):
        self.credentials_path = credentials_path
        self.root_folder_name = root_folder_name
        self.sub_folder_name = sub_folder_name
        self.service = self._authenticate()
        self.target_folder_id = self._prepare_folders()

    def _authenticate(self):
        if not os.path.exists(self.credentials_path):
            raise FileNotFoundError(f"Google Credentials not found at: {self.credentials_path}")
        
        scopes = ['https://www.googleapis.com/auth/drive']
        creds = service_account.Credentials.from_service_account_file(self.credentials_path, scopes=scopes)
        return build('drive', 'v3', credentials=creds)

    def _find_folder(self, name, parent_id=None):
        query = f"mimeType='application/vnd.google-apps.folder' and name='{name}' and trashed=false"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        
        results = self.service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        if files:
            return files[0]['id']
        return None

    def _create_folder(self, name, parent_id=None):
        metadata = {'name': name, 'mimeType': 'application/vnd.google-apps.folder'}
        if parent_id:
            metadata['parents'] = [parent_id]
        file = self.service.files().create(body=metadata, fields='id').execute()
        return file.get('id')

    def _prepare_folders(self):
        # 1. Root
        root_id = self._find_folder(self.root_folder_name)
        if not root_id:
            root_id = self._create_folder(self.root_folder_name)
        
        # 2. Subfolder
        sub_id = self._find_folder(self.sub_folder_name, root_id)
        if not sub_id:
            sub_id = self._create_folder(self.sub_folder_name, root_id)
            
        return sub_id

    def upload_file(self, file_path):
        filename = os.path.basename(file_path)
        metadata = {'name': filename, 'parents': [self.target_folder_id]}
        media = MediaFileUpload(file_path, resumable=True)
        
        file = self.service.files().create(body=metadata, media_body=media, fields='id, webViewLink').execute()
        
        # Make shareable (Anyone with link can view)
        # Note: Depending on organization policy this might be restricted.
        try:
            user_permission = {
                'type': 'anyone',
                'role': 'reader',
            }
            self.service.permissions().create(
                fileId=file.get('id'),
                body=user_permission,
                fields='id',
            ).execute()
        except:
            print("Warning: Could not set public permission. Link might require login.")

        return file.get('webViewLink')

class OneDriveUploader(CloudUploader):
    def __init__(self, config_path, root_folder_name="AutoArchive", sub_folder_name="Scanned"):
        self.config_path = config_path
        self.root_folder_name = root_folder_name
        self.sub_folder_name = sub_folder_name
        self.token = self._authenticate()
        
    def _authenticate(self):
        if not os.path.exists(self.config_path):
            raise FileNotFoundError(f"OneDrive Config not found at: {self.config_path}")
            
        with open(self.config_path, 'r') as f:
            config = json.load(f)
            
        client_id = config.get('client_id')
        client_secret = config.get('client_secret')
        tenant_id = config.get('tenant_id')
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        
        if "access_token" in result:
            return result['access_token']
        else:
            raise Exception(f"Could not acquire OneDrive token: {result.get('error_description')}")

    def upload_file(self, file_path):
        filename = os.path.basename(file_path)
        headers = {'Authorization': 'Bearer ' + self.token}
        
        # 1. Upload File (PUT content)
        # Using "root:/{folder}/{subfolder}/{filename}:/content" syntax
        upload_url = f"https://graph.microsoft.com/v1.0/drive/root:/{self.root_folder_name}/{self.sub_folder_name}/{filename}:/content"
        
        with open(file_path, 'rb') as f:
            data = f.read()
            
        response = requests.put(upload_url, headers=headers, data=data)
        if response.status_code not in [200, 201]:
            raise Exception(f"OneDrive Upload Failed: {response.text}")
            
        item_id = response.json().get('id')
        
        # 2. Create Sharing Link
        link_url = f"https://graph.microsoft.com/v1.0/drive/items/{item_id}/createLink"
        link_payload = {
            "type": "view",
            "scope": "anonymous" # or "organization"
        }
        
        link_res = requests.post(link_url, headers=headers, json=link_payload)
        if link_res.status_code in [200, 201]:
           return link_res.json().get('link', {}).get('webUrl')
        else:
           # Fallback: webUrl from upload response (might require login)
           return response.json().get('webUrl')
