import requests
import pandas as pd
from io import StringIO, BytesIO
from typing import Optional, Dict, List, Any
from pathlib import Path
import logging
import urllib.parse
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed
import time

logger = logging.getLogger(__name__)

class SharePoint():
    def __init__(self, access_token: str, sp_site_id: str):
        self.access_token = access_token
        self.sp_site_id = sp_site_id
        self._drive_id_cache = None

    @property
    def drive_id(self) -> str:
        """
        Resolves the site ID to the canonical drive ID.
        """
        if self._drive_id_cache:
            return self._drive_id_cache

        url = f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive"
        headers = {'Authorization': f'Bearer {self.access_token}'}

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            drive_info = response.json()
            self._drive_id_cache = drive_info.get('id')

            if not self._drive_id_cache:
                raise Exception("Could not find 'id' in /drive response.")

            logger.info(f"Resolved and cached canonical drive ID: {self._drive_id_cache}")
            return self._drive_id_cache
        except Exception as e:
            logger.error(f"Error resolving drive ID: {e}")
            raise

class ReadSP(SharePoint):
    # --- Helper ---
    def get_file_info_from_path(self, path: str) -> Optional[Dict]:
        """Retrieves metadata (including type) for any item (file or folder) at a given path."""
        encoded_path = quote(path, safe='/')
        graph_url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive/root:"
            f"/{encoded_path}"
        )
        headers = {"Authorization": f"Bearer {self.access_token}"}

        try:
            response = requests.get(graph_url, headers=headers)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            logger.error(f"Item not found or inaccessible at path '{path}'. Status: {e.response.status_code}")
            return None
        except Exception as e:
            logger.error(f"Error fetching metadata for path '{path}': {e}")
            return None

    def list_files_in_folder(self, folder_path: str) -> List[Dict]:
        """Lists all CSV files in the specified SharePoint folder and returns their name and ID."""
        encoded_folder_path = quote(folder_path, safe='/')

        current_url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive/root:"
            f"/{encoded_folder_path}:/children"
        )
        headers = {"Authorization": f"Bearer {self.access_token}"}
        all_csv_files = []

        while current_url:
            response = requests.get(current_url, headers=headers)
            response.raise_for_status()
            data = response.json()
            all_csv_files.extend([
                {'name': f['name'], 'id': f['id']}
                for f in data.get("value", [])
                if 'file' in f and f.get('name', '').lower().endswith('.csv')
            ])
            current_url = data.get("@odata.nextLink")
        return all_csv_files

    def get_file_content_by_id(self, item_id: str) -> str:
        """Gets the raw text content of a file using its item ID."""
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive/items/{item_id}/content"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        return response.content.decode('utf-8')

    def get_file_bytes_by_id(self, item_id: str) -> bytes:
        """Gets the raw bytes of a file using its item ID (for binary files e.g. Excel)."""
        graph_url = f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive/items/{item_id}/content"
        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        return response.content

    def list_files_in_folder_by_ext(self, folder_path: str, extensions: List[str]) -> List[Dict]:
        """Lists files matching given extensions in a SharePoint folder."""
        encoded_folder_path = quote(folder_path, safe='/')
        current_url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.sp_site_id}/drive/root:"
            f"/{encoded_folder_path}:/children"
        )
        headers = {"Authorization": f"Bearer {self.access_token}"}
        all_files = []

        while current_url:
            response = requests.get(current_url, headers=headers)
            response.raise_for_status()
            data = response.json()
            all_files.extend([
                {'name': f['name'], 'id': f['id']}
                for f in data.get("value", [])
                if 'file' in f and any(f.get('name', '').lower().endswith(ext) for ext in extensions)
            ])
            current_url = data.get("@odata.nextLink")
        return all_files

    def read_excel(self, path: str, sheet_name=0) -> pd.DataFrame:
        """Read one or all Excel files from a SharePoint path (file or folder)."""
        item_metadata = self.get_file_info_from_path(path)
        if not item_metadata:
            return pd.DataFrame()

        excel_files = []
        if 'file' in item_metadata:
            excel_files = [{'name': item_metadata['name'], 'id': item_metadata['id']}]
        elif 'folder' in item_metadata:
            excel_files = self.list_files_in_folder_by_ext(path, ['.xls', '.xlsx', '.xlsm'])

        excel_files = [f for f in excel_files if not f['name'].startswith('~$')]
        if not excel_files:
            logger.warning(f"No Excel files found at path: {path}")
            return pd.DataFrame()

        def read_single(file_info: Dict) -> Optional[pd.DataFrame]:
            try:
                content = self.get_file_bytes_by_id(file_info['id'])
                df = pd.read_excel(BytesIO(content), sheet_name=sheet_name)
                if df.empty:
                    return pd.DataFrame({'Source.Name': [Path(file_info['name']).stem]})
                df.rename(columns=str, inplace=True)
                df.dropna(how='all', axis=1, inplace=True)
                df['Source.Name'] = Path(file_info['name']).stem
                return df
            except Exception as e:
                logger.warning(f"Could not read {file_info['name']}: {e}")
                return None

        all_dfs = []
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(read_single, f): f['name'] for f in excel_files}
            for future in as_completed(futures):
                df = future.result()
                if df is not None:
                    all_dfs.append(df)

        return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()

    def process_file(self, file_info: Dict) -> Optional[pd.DataFrame]:
        """Downloads, processes, and returns a single DataFrame for a file."""
        try:
            csv_content = self.get_file_content_by_id(file_info['id'])
            df = pd.read_csv(StringIO(csv_content))
            df['source_file'] = file_info['name']
            return df
        except Exception as e:
            logger.error(f"[THREAD] ERROR processing {file_info['name']}: {e}")
            return None

    # --- Execution ---
    def read_csv(self, path: str) -> pd.DataFrame:
        """Master function to orchestrate token acquisition and file download."""
        item_metadata = self.get_file_info_from_path(path)
        if not item_metadata: return pd.DataFrame()

        csv_files = []
        # Single File Mode
        if 'file' in item_metadata:
            csv_files = [{'name': item_metadata['name'], 'id': item_metadata['id']}]
        # Folder Mode
        elif 'folder' in item_metadata:
            csv_files = self.list_files_in_folder(path)

        if not csv_files: return pd.DataFrame()

        all_dataframes = []
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {executor.submit(self.process_file, f): f['name'] for f in csv_files}
            for future in as_completed(futures):
                df = future.result()
                if df is not None: all_dataframes.append(df)

        return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

class UploadSP(SharePoint):
    simple_upload_limit = 4 * 1024 * 1024
    chunk_size = 5 * 1024 * 1024

    @property
    def headers(self) -> Dict[str, str]:
        """Standard headers for most requests."""
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

    # --- Helper ---
    def create_folder_hierarchy(self, folder_path: str) -> str:
        """
        Creates a folder hierarchy using a single Graph API call and returns the ID of the final folder.
        This works because the Graph API supports creating non-existent parent folders when using PUT on a path.
        """
        cleaned_folder_path = folder_path.strip('/')
        encoded_folder_path = urllib.parse.quote(cleaned_folder_path)
        target_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{encoded_folder_path}"

        payload = {
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace"
        }

        response = requests.put(target_url, headers=self.headers, json=payload)
        response.raise_for_status()
        return response.json().get('id')

    def get_folder_id(self, folder_path: str) -> str:
        """
        Gets the Drive Item ID for a folder path using the canonical drive ID.
        If the folder is not found (404), it attempts to create the folder hierarchy.
        """
        cleaned_folder_path = folder_path.strip('/')
        graph_base_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root"

        if not cleaned_folder_path:
            url = graph_base_url
        else:
            encoded_folder_path = urllib.parse.quote(cleaned_folder_path)
            url = f"{graph_base_url}:/{encoded_folder_path}"

        response = requests.get(url, headers=self.headers)

        if response.status_code == 404:
            logger.warning(f"Path {folder_path} not found. Creating...")
            return self.create_folder_hierarchy(cleaned_folder_path)

        response.raise_for_status()
        return response.json().get('id')

    def chunked_upload(self, file_path: Path, upload_url: str) -> Dict[str, Any]:
        """Handles the heavy lifting of splitting and uploading a large file."""
        file_size = file_path.stat().st_size

        session_payload = {'item': {'@microsoft.graph.conflictBehavior': 'replace'}}
        session_resp = requests.post(f"{upload_url}:/createUploadSession", headers=self.headers, json=session_payload)
        session_resp.raise_for_status()
        upload_session_url = session_resp.json().get('uploadUrl')

        with file_path.open('rb') as f:
            start_byte = 0
            while start_byte < file_size:
                chunk = f.read(self.chunk_size)
                end_byte = start_byte + len(chunk) - 1
                chunk_headers = {
                    'Content-Length': str(len(chunk)),
                    'Content-Range': f'bytes {start_byte}-{end_byte}/{file_size}'
                }

                for attempt in range(3):
                    try:
                        response = requests.put(upload_session_url, headers=chunk_headers, data=chunk)
                        if response.status_code in [200, 201]:
                            return response.json()
                        elif response.status_code == 202:
                            break
                        response.raise_for_status()
                    except Exception as e:
                        if attempt == 2: raise
                        time.sleep(2 ** attempt)

                start_byte += len(chunk)
        raise Exception("Upload failed to finalize.")

    # --- Execution ---
    def upload_file(self, original_path: str, destination_path: str, new_name: Optional[str] = None) -> Dict[str, Any]:
        """Main entry point for uploading a file."""
        original_path = Path(original_path)
        if not original_path.exists():
            raise FileNotFoundError(f"File not found at path: {original_path}")

        file_name = new_name or original_path.name
        destination_path = destination_path.replace('\\', '/').strip('/')
        graph_drive_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}"

        if original_path.stat().st_size <= self.simple_upload_limit:
            # Simple upload
            target_path = f"{destination_path}/{file_name}".strip('/')
            url = f"{graph_drive_url}/root:/{urllib.parse.quote(target_path)}:/content"
            headers = {**self.headers, 'Content-Type': 'application/octet-stream'}

            with original_path.open('rb') as f:
                response = requests.put(url, headers=headers, data=f.read())
            response.raise_for_status()
            return response.json()

        else:
            # Chunked upload
            folder_id = self.get_folder_id(destination_path)
            upload_url = f"{graph_drive_url}/items/{folder_id}:/{file_name}"

            # Delete existing item first to clear any stale upload sessions (avoids 409 Conflict)
            item_url = f"{graph_drive_url}/items/{folder_id}:/{urllib.parse.quote(file_name)}"
            requests.delete(item_url, headers=self.headers)  # 404 is fine if it doesn't exist yet

            return self.chunked_upload(original_path, upload_url)
