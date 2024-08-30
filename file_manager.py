import os
import tempfile
import sys
import shutil

class FileManager:
    """
    A class to manage file operations such as creating folders and deleting files.
    """

    @staticmethod
    def create_folder_desktop(folder_name):
        """
        Create a folder on the desktop of the user.

        Parameters:
        folder_name (str): The name of the folder to be created.

        Returns:
        str: The path to the created folder.
        """
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        folder_path = os.path.join(desktop_path, folder_name)
        os.makedirs(folder_path, exist_ok=True)
        return folder_path

    @staticmethod
    def delete_file(file_path):
        """
        Delete a file at a specified filepath.

        Parameters:
        file_path (str): The path to the file to be deleted.

        Returns:
        bool: True if the file was deleted, False otherwise.
        """
        try:
            os.remove(file_path)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            print(f"Error deleting file: {e}")
            return False

    @staticmethod
    def resource_path(relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
