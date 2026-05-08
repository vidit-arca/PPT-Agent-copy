import io
import logging
from minio import Minio
from src.config.settings import get_config

class MinioClient:
    """Client for interacting with MinIO object storage."""
    
    def __init__(self):
        config = get_config()
        if not config.is_minio_configured():
            raise ValueError("MinIO is not configured. Please check environment variables.")
            
        self.client = Minio(
            config.minio_endpoint,
            access_key=config.minio_access_key,
            secret_key=config.minio_secret_key,
            secure=config.minio_secure
        )
        self.bucket_name = config.minio_bucket_name
        
        # Verify connection
        try:
            if not self.client.bucket_exists(self.bucket_name):
                logging.warning(f"Bucket '{self.bucket_name}' does not exist. Attempting to create it...")
                self.client.make_bucket(self.bucket_name)
                logging.info(f"Created bucket '{self.bucket_name}'")
        except Exception as e:
            logging.error(f"Failed to connect to MinIO: {e}")
            raise

    def list_excel_files(self, prefix: str = "") -> list[str]:
        """
        Recursively list all Excel files (.xls, .xlsx) in the bucket, optionally filtering by prefix.
        
        Args:
            prefix: Optional folder path prefix to filter files (e.g., "weekly_outputs/2025-01-01/")
            
        Returns:
            List of object names.
        """
        excel_files = []
        try:
            # recursive=True allows traversing all subfolders
            objects = self.client.list_objects(self.bucket_name, prefix=prefix, recursive=True)
            
            for obj in objects:
                if obj.object_name.lower().endswith(('.xls', '.xlsx')):
                    excel_files.append(obj.object_name)
                    
            logging.info(f"Found {len(excel_files)} Excel files in MinIO bucket '{self.bucket_name}' with prefix '{prefix}'")
            return excel_files
        except Exception as e:
            logging.error(f"Error listing files from MinIO: {e}")
            return []

    def get_file_content(self, object_name: str) -> io.BytesIO:
        """
        Download file content from MinIO into memory.
        Returns BytesIO object containing file data.
        """
        try:
            logging.info(f"Downloading '{object_name}' from MinIO...")
            response = self.client.get_object(self.bucket_name, object_name)
            file_data = io.BytesIO(response.read())
            response.close()
            response.release_conn()
            return file_data
        except Exception as e:
            logging.error(f"Error downloading file '{object_name}': {e}")
            raise
