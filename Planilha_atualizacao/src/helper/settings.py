import json
from typing import Dict

class Settings:
    """Class to manage configuration settings from a JSON file."""

    def __init__(self, config_path: str = "input"):
        """
        Initialize Settings with the path to the configuration file.

        Args:
            config_path (str): The path to the configuration JSON file.
        """
        self.config_path = config_path
        self._load_config()

    def _load_config(self) -> None:
        """Loads the configuration from the JSON file and sets instance variables."""
        with open(self.config_path, 'r') as file:
            config = json.load(file)

            self.input_folder_path = config.get('input_folder_path', '')
            self.log_file_path = config.get('log_file_path', '')

            messages_config = config.get('messages', {})

            self.valid_local_list = messages_config.get('valid_local_list', '')
            self.valid_status_list = messages_config.get('valid_status_list', '')
            self.valid_bool_list = messages_config.get('valid_bool_list', '')
            self.valid_sex_list = messages_config.get('valid_sex_list', '')
            self.valid_returned_list = messages_config.get('valid_returned_list', '')
            self.valid_race_list = messages_config.get('valid_race_list', '')
            self.valid_color_list = messages_config.get('valid_color_list', '')
            self.valid_fiv_felv_list = messages_config.get('valid_fiv_felv_list', '')
            self.valid_vaccine_list = messages_config.get('valid_vaccine_list', '')
            self.valid_profile_list = messages_config.get('valid_profile_list', '')

            
            
            self.trigger_word = messages_config.get('trigger_word', '')
            self.ending_word = messages_config.get('ending_word', '')

            self.local_col = messages_config.get('local_col', '')
            self.status_col = messages_config.get('status_col', '')
            self.azure_col = messages_config.get('azure_col', '')
            self.simplesvet_col = messages_config.get('simplesvet_col', '')
            self.name_col = messages_config.get('name_col', '')
            self.ngo_entry_col = messages_config.get('ngo_entry_col', '')
            self.ngo_exit_col = messages_config.get('ngo_exit_col', '')
            self.vaccination_card_col = messages_config.get('vaccination_card_col', '')
            self.birthday_col = messages_config.get('birthday_col', '')
            self.sex_col = messages_config.get('sex_col', '')
            self.race_col = messages_config.get('race_col', '')
            self.color_col = messages_config.get('color_col', '')
            self.date_neuter_col = messages_config.get('date_neuter_col', '')
            self.FIV_col = messages_config.get('FIV_col', '')
            self.FELV_col = messages_config.get('FELV_col', '')
            self.date_FIV_FELV_col = messages_config.get('date_FIV_FELV_col', '')
            self.date_2nd_dose_vaccine_col = messages_config.get('date_2nd_dose_vaccine_col', '')
            self.vaccine_type_col = messages_config.get('vaccine_type_col', '')
            self.date_vaccine_renew_col = messages_config.get('date_vaccine_renew_col', '')
            self.date_rabies_vaccine_col = messages_config.get('date_rabies_vaccine_col', '')
            self.date_rabies_renew_col = messages_config.get('date_rabies_renew_col', '')
            self.history_col = messages_config.get('history_col', '')
            self.family_col = messages_config.get('family_col', '')
            self.notes_col = messages_config.get('notes_col', '')
            self.microchip_col = messages_config.get('microchip_col', '')
            self.interaction_animals_col = messages_config.get('interaction_animals_col', '')
            self.interaction_humans_col = messages_config.get('interaction_humans_col', '')
            self.profile_col = messages_config.get('profile_col', '')
            self.returned_col = messages_config.get('returned_col', '')
            self.date_returned_col = messages_config.get('date_returned_col', '')
            self.reason_returned_col = messages_config.get('reason_returned_col', '')
            self.health_notes_col = messages_config.get('health_notes_col', '')
            self.admin_notes_col = messages_config.get('admin_notes_col', '')


            sheets_config = config.get('google_sheets', {})
            self.credentials_path = sheets_config.get('credentials_path', '')
            self.google_sheets_id = sheets_config.get('id', '')
            self.google_sheets_tab = sheets_config.get('tab', '')
