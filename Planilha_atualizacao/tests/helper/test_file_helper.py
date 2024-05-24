import os
import json
import pytest
from src.helper.file_helper import FileHelper
from src.helper.settings import Settings

@pytest.fixture
def setup_files_and_config(tmp_path):
    # Create temporary input folder and files for testing
    folder = tmp_path / "input"
    folder.mkdir()
    file1 = folder / "message1.txt"
    file1.write_text("This is message 1")
    file2 = folder / "message2.txt"
    file2.write_text("This is message 2")
    
    # Create a temporary config folder and config file
    config_folder = tmp_path / "config"
    config_folder.mkdir()
    config_path = config_folder / "settings.json"
    
    config_data = {
        "input_folder_path": str(folder),
        "messages": {
            "trigger_word": "UPDATE",
            "column_name": "Status",
            "column_id": "A"
        },
        "google_sheets": {
            "link": "https://docs.google.com/spreadsheets/d/your_sheet_id",
            "tab": "Sheet1",
            "username": "your_username",
            "password": "your_password"
        }
    }
    
    with open(config_path, 'w') as config_file:
        json.dump(config_data, config_file)
    
    return config_path

def test_read_files(setup_files_and_config):
    config_path = setup_files_and_config
    settings = Settings(str(config_path))
    file_helper = FileHelper(settings)
    file_helper.read_files()
    assert file_helper.get_contents() == ["This is message 1", "This is message 2"]

def test_remove_files(setup_files_and_config):
    config_path = setup_files_and_config
    settings = Settings(str(config_path))
    file_helper = FileHelper(settings)
    file_helper.read_files()
    file_helper.remove_files()
    input_folder = settings.input_folder_path
    assert len(os.listdir(input_folder)) == 0

def test_get_contents(setup_files_and_config):
    config_path = setup_files_and_config
    settings = Settings(str(config_path))
    file_helper = FileHelper(settings)
    assert file_helper.get_contents() == []
    file_helper.read_files()
    assert file_helper.get_contents() == ["This is message 1", "This is message 2"]
