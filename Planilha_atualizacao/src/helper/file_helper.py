import markdown
import os
import webbrowser

from typing import List
from .settings import Settings

class FileHelper:
    """Helper class to manage file operations in the input folder."""

    def __init__(self, settings: Settings):
        """
        Initialize FileHelper with the Settings instance.

        Args:
            settings (Settings): An instance of the Settings class.
        """
        self.folder_path = settings.input_folder_path
        self.contents = []

    def read_files(self, folder_path: str = None) -> None:
        """Reads all text files in the folder and stores their contents in a variable."""
        self.contents.clear()
        if not folder_path:
            folder_path = self.folder_path
        for filename in os.listdir(path=folder_path):
            if filename.endswith(".txt"):
                with open(os.path.join(folder_path, filename), 'r') as file:
                    self.contents.append(file.read())

    def remove_files(self) -> None:
        """Removes all text files in the folder."""
        for filename in os.listdir(self.folder_path):
            if filename.endswith(".txt"):
                os.remove(os.path.join(self.folder_path, filename))
    @staticmethod
    def create_empty_txt_file(self):
        """
        Create an empty text file.

        Args:
            file_path (str): Path to the text file to create.
        """
        
        with open(os.path.join(self.folder_path, "mensagens.txt"), 'w') as file:
            pass  # Create an empty file

    def get_contents(self) -> List[str]:
        """
        Get the contents read from the files.

        Returns:
            List[str]: List of strings containing the contents of each file.
        """
        return self.contents

    @staticmethod
    def append_to_file(file_path: str, content: str) -> None:
        """
        Append content to a file.

        Args:
            file_path (str): The path to the file.
            content (str): The content to append.
        """
        with open(file_path, 'a') as file:
            file.write(content + '\n')
    
    @staticmethod
    def remove_file_if_exists(file_path: str):
        """
        Remove a file if it exists.

        Args:
            file_path (str): The path to the file to remove.
        """
        if os.path.exists(file_path):
            os.remove(file_path)

    def read_file(self, file_path: str) -> List[str]:
        """
        Read content from a file.

        Args:
            file_path (str): The path to the file.

        Returns:
            List[str]: List of strings containing each line of the file.
        """
        with open(file_path, 'r') as file:
            return file.readlines()

    @staticmethod
    def convert_md_to_html(md_file_path: str, html_file_path: str):
        """
        Convert a markdown file to an HTML file.

        Args:
            md_file_path (str): Path to the markdown file.
            html_file_path (str): Path to the output HTML file.
        """
        # Read the markdown file
        with open(md_file_path, 'r', encoding='utf-8') as md_file:
            md_content = md_file.read()

        # Convert markdown content to HTML
        html_content = markdown.markdown(md_content)

        # Write the HTML content to a file
        with open(html_file_path, 'w', encoding='utf-8') as html_file:
            html_file.write(html_content)

        print(f"Markdown file has been converted to HTML and saved as {html_file_path}")

    @staticmethod
    def open_file_in_browser(url: str):
        webbrowser.open(url = url)
