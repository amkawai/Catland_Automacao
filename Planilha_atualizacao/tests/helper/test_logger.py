import json
import pytest
import os
from src.helper.settings import Settings
from src.helper.logger import Logger, LogLevel

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

@pytest.fixture
def logger(setup_config, tmp_path):
    config_path = setup_config
    settings = Settings(str(config_path))
    logger = Logger(settings)
    logger.log_file_path = tmp_path / "log.txt"  # Use temporary log file
    return logger

def test_log_info(logger):
    logger.log(LogLevel.INFO, "This is an info message.")
    logs = logger.read_logs()
    assert len(logs) == 1
    assert "INFO" in logs[0]
    assert "This is an info message." in logs[0]

def test_log_warning(logger):
    logger.log(LogLevel.WARNING, "This is a warning message.")
    logs = logger.read_logs()
    assert len(logs) == 1
    assert "WARNING" in logs[0]
    assert "This is a warning message." in logs[0]

def test_log_error(logger):
    logger.log(LogLevel.ERROR, "This is an error message.")
    logs = logger.read_logs()
    assert len(logs) == 1
    assert "ERROR" in logs[0]
    assert "This is an error message." in logs[0]

def test_multiple_logs(logger):
    logger.log(LogLevel.INFO, "First message.")
    logger.log(LogLevel.WARNING, "Second message.")
    logger.log(LogLevel.ERROR, "Third message.")
    logs = logger.read_logs()
    assert len(logs) == 3
    assert "INFO" in logs[0]
    assert "First message." in logs[0]
    assert "WARNING" in logs[1]
    assert "Second message." in logs[1]
    assert "ERROR" in logs[2]
    assert "Third message." in logs[2]
