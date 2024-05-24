import json
import pytest
from src.helper.settings import Settings

@pytest.fixture
def setup_config(tmp_path):
    # Create a temporary config folder and config file
    config_folder = tmp_path / "config"
    config_folder.mkdir()
    config_path = config_folder / "settings.json"
    
    config_data = {
        "input_folder_path": "input",
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

def test_settings_initialization(setup_config):
    config_path = setup_config
    settings = Settings(str(config_path))
    
    assert settings.input_folder_path == "input"
    assert settings.trigger_word == "UPDATE"
    assert settings.column_name == "Status"
    assert settings.column_id == "A"
    assert settings.google_sheets_link == "https://docs.google.com/spreadsheets/d/your_sheet_id"
    assert settings.google_sheets_tab == "Sheet1"
    assert settings.google_sheets_username == "your_username"
    assert settings.google_sheets_password == "your_password"

def test_settings_missing_values(tmp_path):
    # Create a temporary config folder and config file with missing values
    config_folder = tmp_path / "config"
    config_folder.mkdir()
    config_path = config_folder / "settings.json"
    
    config_data = {
        "input_folder_path": "input",
        "messages": {
            "trigger_word": "UPDATE"
        }
    }
    
    with open(config_path, 'w') as config_file:
        json.dump(config_data, config_file)
    
    settings = Settings(str(config_path))
    
    assert settings.input_folder_path == "input"
    assert settings.trigger_word == "UPDATE"
    assert settings.column_name == ""
    assert settings.column_id == ""
    assert settings.google_sheets_link == ""
    assert settings.google_sheets_tab == ""
    assert settings.google_sheets_username == ""
    assert settings.google_sheets_password == ""
