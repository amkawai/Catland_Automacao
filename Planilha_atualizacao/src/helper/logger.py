from datetime import datetime
from enum import Enum
from typing import List
from .file_helper import FileHelper
from .settings import Settings

settings = Settings("config/settings.json")

class Logger:
    """Logger class to handle logging messages with different levels and timestamps."""
    
    @staticmethod
    def log(level: str, message: str) -> None:
        """
        Log a message with a given level and timestamp.

        Args:
            level (LogLevel): The level of the log message.
            message (str): The log message.
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"{timestamp} - {level.upper()} - {message}"
        print(log_message)  # Print the log message to console
        if level == "ERROR":
            #settings = Settings("config/settings.json")
            #FileHelper.append_to_file(file_path=settings.log_file_path, content=f"{message}")
            Logger.add_horizontal_line()
            Logger.write_to_log_file(message = f"{message}")

    
    @staticmethod
    def add_horizontal_line():
        Logger.write_to_log_file('\n---\n')

    @staticmethod
    def initialize_log():
        Logger.write_to_log_file('#Erros identificados\n')
        Logger.write_to_log_file('Início:')
        Logger.write_timestamp()
        Logger.write_to_log_file('\n\n')

    @staticmethod
    def write_to_log_file(message: str) -> None:
        FileHelper.append_to_file(file_path=settings.log_file_path, content=f'\n{message}\n')
    
    @staticmethod
    def write_timestamp() -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        Logger.write_to_log_file(f'\n{timestamp}')
    