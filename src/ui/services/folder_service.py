"""
EchoScribe AI - Folder Service Module
Handles folder operations including open folder functionality
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Optional, List, Dict, Any
import json
from pathlib import Path

class FolderService:
    """Service for managing folder operations and file organization"""

    def __init__(self, config_manager=None):
        """Initialize folder service with optional config manager"""
        self.config_manager = config_manager
        self.recent_folders = []
        self.load_recent_folders()

    def load_recent_folders(self) -> None:
        """Load recent folders from configuration"""
        try:
            if self.config_manager:
                recent = self.config_manager.get('recent_folders', [])
                self.recent_folders = recent[-10:]  # Keep last 10
            else:
                self.recent_folders = []
        except Exception:
            self.recent_folders = []

    def save_recent_folders(self) -> None:
        """Save recent folders to configuration"""
        try:
            if self.config_manager:
                self.config_manager.set('recent_folders', self.recent_folders)
        except Exception:
            pass

    def open_folder_dialog(self, parent_window=None) -> Optional[str]:
        """
        Open a folder selection dialog

        Args:
            parent_window: Parent tkinter window

        Returns:
            Selected folder path or None if cancelled
        """
        try:
            # Create root window if not provided
            if parent_window is None:
                root = tk.Tk()
                root.withdraw()
            else:
                root = parent_window

            # Configure dialog options
            options = {
                'title': 'Select Folder for EchoScribe AI',
                'mustexist': True,
                'initialdir': self.get_last_folder() or os.path.expanduser('~')
            }

            # Show folder dialog
            folder_path = filedialog.askdirectory(**options)

            # Clean up temporary root if created
            if parent_window is None and 'root' in locals():
                root.destroy()

            if folder_path:
                self.add_recent_folder(folder_path)
                return folder_path

            return None

        except Exception as e:
            messagebox.showerror("Error", f"Failed to open folder dialog: {str(e)}")
            return None

    def add_recent_folder(self, folder_path: str) -> None:
        """Add folder to recent folders list"""
        if folder_path and os.path.isdir(folder_path):
            if folder_path in self.recent_folders:
                self.recent_folders.remove(folder_path)
            self.recent_folders.insert(0, folder_path)
            self.recent_folders = self.recent_folders[:10]  # Keep max 10
            self.save_recent_folders()

    def get_last_folder(self) -> Optional[str]:
        """Get the last used folder"""
        return self.recent_folders[0] if self.recent_folders else None

    def get_recent_folders(self) -> List[str]:
        """Get list of recent folders"""
        return [f for f in self.recent_folders if os.path.isdir(f)]

    def create_folder_structure(self, base_path: str, structure: Dict[str, Any]) -> bool:
        """
        Create folder structure based on configuration

        Args:
            base_path: Base directory path
            structure: Dictionary defining folder structure

        Returns:
            True if successful, False otherwise
        """
        try:
            base_path = Path(base_path)
            base_path.mkdir(parents=True, exist_ok=True)

            for folder_name, sub_structure in structure.items():
                folder_path = base_path / folder_name
                folder_path.mkdir(exist_ok=True)

                if isinstance(sub_structure, dict):
                    self.create_folder_structure(str(folder_path), sub_structure)

            return True

        except Exception as e:
            print(f"Error creating folder structure: {e}")
            return False

    def get_folder_contents(self, folder_path: str) -> Dict[str, List[str]]:
        """
        Get contents of a folder

        Args:
            folder_path: Path to folder

        Returns:
            Dictionary with files and subfolders
        """
        try:
            folder = Path(folder_path)
            if not folder.is_dir():
                return {"files": [], "folders": []}

            files = [f.name for f in folder.iterdir() if f.is_file()]
            folders = [f.name for f in folder.iterdir() if f.is_dir()]

            return {"files": sorted(files), "folders": sorted(folders)}

        except Exception as e:
            print(f"Error getting folder contents: {e}")
            return {"files": [], "folders": []}

    def validate_folder(self, folder_path: str) -> tuple[bool, str]:
        """
        Validate if folder exists and is accessible

        Args:
            folder_path: Path to validate

        Returns:
            Tuple of (is_valid, error_message)
        """
        if not folder_path:
            return False, "No folder path provided"

        if not os.path.exists(folder_path):
            return False, f"Folder does not exist: {folder_path}"

        if not os.path.isdir(folder_path):
            return False, f"Path is not a directory: {folder_path}"

        if not os.access(folder_path, os.R_OK):
            return False, f"Cannot read folder: {folder_path}"

        return True, ""

    def get_folder_size(self, folder_path: str) -> int:
        """Get total size of folder in bytes"""
        try:
            total_size = 0
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for filename in filenames:
                    filepath = os.path.join(dirpath, filename)
                    if os.path.exists(filepath):
                        total_size += os.path.getsize(filepath)
            return total_size
        except Exception:
            return 0

    def get_folder_info(self, folder_path: str) -> Dict[str, Any]:
        """Get comprehensive folder information"""
        try:
            folder = Path(folder_path)
            if not folder.is_dir():
                return {}

            stats = folder.stat()
            contents = self.get_folder_contents(folder_path)

            return {
                "path": str(folder),
                "name": folder.name,
                "size": self.get_folder_size(folder_path),
                "file_count": len(contents["files"]),
                "folder_count": len(contents["folders"]),
                "created": stats.st_ctime,
                "modified": stats.st_mtime,
                "accessible": os.access(folder_path, os.R_OK | os.W_OK)
            }

        except Exception as e:
            return {"error": str(e)}

# Global instance for easy access
folder_service = FolderService()
