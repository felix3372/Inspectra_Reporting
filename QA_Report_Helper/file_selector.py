"""
File selector module for hierarchical network path file selection.
Allows users to navigate Month → Campaign → Excel File structure.
"""

import os
import logging
from pathlib import Path
from typing import List, Optional, Tuple, Dict
from datetime import datetime
from utils.config import BASE_DIR


logger = logging.getLogger(__name__)


class FileSelector:
    """Handles hierarchical file selection from network paths (Month → Campaign → File)."""
    
    # Base directory for reports - configure this to your network path
    BASE_DIR
    
    def __init__(self, base_dir: Optional[str] = None):
        """
        Initialize FileSelector with base directory.
        
        Args:
            base_dir: Base directory path. If None, uses BASE_DIR
        """
        self.base_dir = base_dir or self.BASE_DIR
        self.supported_extensions = ['.xlsx', '.xlsm']
    
    def path_exists(self, path: str = None) -> bool:
        """Check if path exists and is accessible."""
        check_path = path or self.base_dir
        try:
            return os.path.exists(check_path) and os.path.isdir(check_path)
        except Exception as e:
            logger.error(f"Error checking path: {str(e)}")
            return False
    
    def get_month_folders(self) -> List[str]:
        """
        Get list of month folders from base directory.
        
        Returns:
            List of month folder names (e.g., ["April'25", "May'25"])
        """
        months = []
        
        if not self.path_exists():
            logger.warning(f"Base directory not accessible: {self.base_dir}")
            return months
        
        try:
            for item in os.listdir(self.base_dir):
                item_path = os.path.join(self.base_dir, item)
                if os.path.isdir(item_path):
                    months.append(item)
            
            # Sort months (you can customize sorting logic)
            months.sort()
            
            logger.info(f"Found {len(months)} month folders")
            return months
            
        except Exception as e:
            logger.error(f"Error reading month folders: {str(e)}")
            return []
    
    def get_campaign_folders(self, month: str) -> List[str]:
        """
        Get list of campaign folders for a specific month.
        
        Args:
            month: Month folder name (e.g., "April'25")
            
        Returns:
            List of campaign folder names (e.g., ["6326_Apr'25", "6327_Apr'25"])
        """
        campaigns = []
        month_path = os.path.join(self.base_dir, month)
        
        if not self.path_exists(month_path):
            logger.warning(f"Month folder not accessible: {month_path}")
            return campaigns
        
        try:
            for item in os.listdir(month_path):
                item_path = os.path.join(month_path, item)
                if os.path.isdir(item_path):
                    campaigns.append(item)
            
            # Sort campaigns
            campaigns.sort()
            
            logger.info(f"Found {len(campaigns)} campaign folders in {month}")
            return campaigns
            
        except Exception as e:
            logger.error(f"Error reading campaign folders: {str(e)}")
            return []
    
    def get_excel_files(self, month: str, campaign: str) -> List[Tuple[str, str, datetime, float]]:
        """
        Get list of Excel files for a specific campaign.
        
        Args:
            month: Month folder name
            campaign: Campaign folder name
            
        Returns:
            List of tuples: (filename, full_path, modified_time, size_mb)
        """
        files = []
        campaign_path = os.path.join(self.base_dir, month, campaign)
        
        if not self.path_exists(campaign_path):
            logger.warning(f"Campaign folder not accessible: {campaign_path}")
            return files
        
        try:
            for file in os.listdir(campaign_path):
                file_path = os.path.join(campaign_path, file)
                
                # Check if it's a file and has correct extension
                if os.path.isfile(file_path):
                    _, ext = os.path.splitext(file)
                    if ext.lower() in self.supported_extensions:
                        # Get file metadata
                        mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                        size_bytes = os.path.getsize(file_path)
                        size_mb = size_bytes / (1024 * 1024)
                        
                        files.append((file, file_path, mod_time, size_mb))
            
            # Sort by modification time (newest first)
            files.sort(key=lambda x: x[2], reverse=True)
            
            logger.info(f"Found {len(files)} Excel files in {campaign}")
            return files
            
        except Exception as e:
            logger.error(f"Error reading Excel files: {str(e)}")
            return []
    
    def get_file_display_name(self, filename: str, mod_time: datetime, size_mb: float) -> str:
        """
        Create user-friendly display name for a file.
        
        Args:
            filename: Name of the file
            mod_time: File modification time
            size_mb: File size in MB
            
        Returns:
            Formatted display name
        """
        return f"{filename} ({size_mb:.1f} MB, {mod_time.strftime('%d-%b-%Y %I:%M %p')})"
    
    def read_file(self, file_path: str) -> Optional[bytes]:
        """
        Read file content from network path.
        
        Args:
            file_path: Full path to the file
            
        Returns:
            File content as bytes, or None if error
        """
        try:
            with open(file_path, 'rb') as f:
                return f.read()
        except Exception as e:
            logger.error(f"Error reading file {file_path}: {str(e)}")
            return None
    
    def validate_file_access(self, file_path: str) -> Tuple[bool, str]:
        """
        Validate that file exists and is readable.
        
        Args:
            file_path: Full path to the file
            
        Returns:
            Tuple of (is_valid, message)
        """
        if not os.path.exists(file_path):
            return False, "File does not exist"
        
        if not os.path.isfile(file_path):
            return False, "Path is not a file"
        
        try:
            # Try to open file to check permissions
            with open(file_path, 'rb') as f:
                f.read(1)
            return True, "File is accessible"
        except PermissionError:
            return False, "Permission denied - check file access rights"
        except Exception as e:
            return False, f"Error accessing file: {str(e)}"
    
    def get_full_path(self, month: str, campaign: str, filename: str) -> str:
        """
        Construct full file path from components.
        
        Args:
            month: Month folder name
            campaign: Campaign folder name
            filename: File name
            
        Returns:
            Full path to file
        """
        return os.path.join(self.base_dir, month, campaign, filename)