from typing import Dict, List
from .settings import Settings

class Updater:
    """Class to parse updates from text files."""

    def __init__(self, settings: Settings):
        """
        Initialize Updater with the Settings instance.

        Args:
            settings (Settings): An instance of the Settings class.
        """
        self.settings = settings

    def parse_updates(self, file_contents: str) -> List[Dict[str, str]]:
        """
        Parse updates from file contents between trigger and ending words.

        Args:
            file_contents (str): Contents of the text file.

        Returns:
            List[Dict[str, str]]: List of dictionaries containing parsed updates.
        """
        trigger_word = self.settings.trigger_word
        ending_word = self.settings.ending_word
        updates = []
        start_index = file_contents.find(trigger_word)
        while start_index != -1:
            end_index = file_contents.find(ending_word, start_index + len(trigger_word))
            if end_index != -1:
                update_data = self._parse_message(file_contents[start_index:end_index])
                if update_data:
                    updates.append(update_data)
                start_index = file_contents.find(trigger_word, end_index)
            else:
                break
        return updates

    def _parse_message(self, message: str) -> Dict[str, str]:
        """
        Parse a single message and extract relevant data into a dictionary.

        Args:
            message (str): The message to parse.

        Returns:
            Dict[str, str]: A dictionary containing parsed data.
        """
        lines = message.split('\n')
        data = {}
        for line in lines[1:-1]:
            try:
                if len(line.split(":")) == 2:
                    data[line.split(":")[0].strip()] = line.split(":")[1].strip()
                if len(line.split(":")) > 2:
                    temp_data = line
                    temp_data = temp_data.replace(line.split(":")[0].strip(),"")
                    temp_data = temp_data[2:]
                    data[line.split(":")[0].strip()] = temp_data
            except:
                pass
        data["Mensagem Original"] = message.replace("\n", "\n\n")
        return data
    