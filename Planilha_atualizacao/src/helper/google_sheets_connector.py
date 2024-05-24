import pygsheets
from typing import List, Dict

class GoogleSheetsConnector:
    """Class to connect to and interact with Google Sheets using OAuth2 flow and credentials file."""

    def __init__(self, credentials_path: str, spreadsheet_id: str, sheet_name: str):
        """
        Initialize GoogleSheetsConnector and authenticate with Google Sheets.

        Args:
            credentials_path (str): Path to the Google Sheets credentials JSON file.
            spreadsheet_id (str): ID of the Google Sheets spreadsheet.
            sheet_name (str): Name of the sheet within the spreadsheet.
        """
        self.credentials_path = credentials_path
        self.spreadsheet_id = spreadsheet_id
        self.sheet_name = sheet_name
        self.client = None
        self.sheet = None
        self._connect()

    def _connect(self):
        """Authenticate and connect to Google Sheets."""
        self.client = pygsheets.authorize(service_file=self.credentials_path)
        self.sheet = self.client.open_by_key(self.spreadsheet_id).worksheet_by_title(self.sheet_name)

    def validate_cat_data(self, cat_id: str, cat_name: str, id_col: str = "D", name_col: str = "E") -> bool:
        """
        Validate cat data by checking if the cat name matches the name in the spreadsheet.

        Args:
            cat_id (str): The ID of the cat to validate.
            cat_name (str): The name of the cat to validate.
            id_col (str): The column containing the cat IDs in the spreadsheet.
            name_col (str): The column containing the cat names in the spreadsheet.

        Returns:
            bool: True if the cat name matches the name in the spreadsheet, False otherwise.
        """
        id_col_number = self.column_letter_to_number(column_letter=id_col)
        name_col_number = self.column_letter_to_number(column_letter=name_col)

        try:
            id_row = self.get_row_from_data(pattern = cat_id, col = id_col)
            name_row = self.get_row_from_data(pattern=cat_name, col = name_col)
            if id_row == name_row:
                return True
            return False
        except Exception as e:
            #print(e)
            return False
            
    @staticmethod
    def column_letter_to_number(column_letter: str) -> int:
        """
        Convert a column letter to its corresponding number.

        Args:
            column_letter (str): The column letter to convert (e.g., "A", "B", "C", ...).

        Returns:
            int: The corresponding column number (1-based index).
        """
        column_number = 0
        for char in column_letter:
            column_number = column_number * 26 + (ord(char.upper()) - ord('A')) + 1
        return column_number

    def update_value(self, row_number: int, col_letter: str, value: any) -> None:
        """ Replaces the value in a cell

        Args:
            row_number (int): number of the row to update
            col_letter (str): column_letter(s)
            value (any): value to insert in cell
        """
        col_number = GoogleSheetsConnector.column_letter_to_number(column_letter=col_letter)
        self.sheet.update_value(addr = (row_number, col_number), val = value, parse = True)

    def get_row_from_data(self, pattern: any, matchCase: bool = True, matchEntireCell: bool = True, col: str = None) -> int:
        if not col:
            return self.sheet.find(pattern = pattern, matchCase = matchCase, matchEntireCell = matchEntireCell)[0].row    
        col_number = GoogleSheetsConnector.column_letter_to_number(column_letter = col)
        return self.sheet.find(pattern = pattern, matchCase = matchCase, matchEntireCell = matchEntireCell, cols = (col_number, col_number))[0].row
    
    def get_data_from_cell(self, row_number: int, col_letter: str) -> any:
        col_number = GoogleSheetsConnector.column_letter_to_number(column_letter=col_letter)
        return self.sheet.cell(addr = (row_number, col_number))
    
    def append_data_to_cell(self, row_number: int, col_letter: str, additional_value: any) -> None:
        current_data = self.get_data_from_cell(row_number = row_number, col_letter = col_letter).value
        value = str(current_data) + "\n" + str(additional_value)
        self.update_value(row_number = row_number, col_letter = col_letter, value = value)

