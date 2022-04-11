# Google Sheets API Doc: https://developers.google.com/sheets/api/reference/rest


from googleapiclient.discovery import build, Resource
from google.oauth2 import service_account
from typing import Optional

class GoogleSheetsTool:
    """Wrapper API for the Google API."""

    def __init__(self):

        self.filename = ""
        self.spreadsheet_id = ""
        self.folder_id = ""
        self.current_sheets: dict = {"Sheet1":{"id":0, "next_row":1}}
        self.sheet_id_runner: int = 1
        self.requests = list()
        self.update_values_requests = list()
        self.service: Optional[Resource] = None
        self.sheet: Optional[Resource] = None


    def authenticate(self, service_account_file) -> None:
        """Athenticate and connect to Google services for spread sheets."""
        
        g_scopes: list = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_file(
            service_account_file,
            scopes=g_scopes
        )

        self.service = build("sheets", "v4", credentials=creds)
        self.sheet = self.service.spreadsheets()

        response: dict = self.build_spread_sheet(creds)
        self.spreadsheet_id = response["id"]


    def build_spread_sheet(self, creds) -> dict:
        """Create a google sheet in the "parents" folder. """

        drive = build('drive', 'v3', credentials=creds)
        file_metadata = {
            "name": self.filename,
            "parents": [self.folder_id],
            "mimeType": "application/vnd.google-apps.spreadsheet"
        }
        return drive.files().create(body=file_metadata).execute()


    def get_values_by_range(self, cell_range: str) -> Optional[dict]:
        """Returns the values of a specified range."""
        response = self.sheet.values().get(spreadsheetId=self.spreadsheet_id, range=cell_range).execute()
        return response.get("values")


    def get_next_row(self, sheet_name: str) -> int:
        """Returns the last row number that has been appended to the specified sheet."""

        return self.current_sheets[sheet_name]["next_row"]


    def get_sheet_properties(self) -> dict:
        """Return the spreadsheet properties as a dict.

        Keys on returned dict:
            - "properties" - General properties of the entire spreadsheet.
            - "sheets" - Information on all current sheets in the spreadsheet.
            - "spreadsheetUrl"
        """

        return self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()


    def set_file_and_folder_info(self, file_and_folder_info: tuple) -> None:
        """Set the filename and folder information for the sheet export."""

        self.filename = file_and_folder_info[0]
        self.folder_id = file_and_folder_info[1]

    
    def values_batch_update(self) -> None:
        """Batch updates all current value requests."""

        values_body = {"valueInputOption": "USER_ENTERED", "data": self.update_values_requests}
        self.sheet.values().batchUpdate(spreadsheetId=self.spreadsheet_id, body=values_body).execute()
        self.update_values_requests.clear()


    def batch_update(self) -> None:
        """Batch updates all current requests."""
        
        body: dict = {"requests": self.requests}        
        self.sheet.batchUpdate(spreadsheetId=self.spreadsheet_id, body=body).execute()
        self.requests.clear()
    
    def add_values_request(self, cell_range: str, rows: list) -> None:
        """Add values in a range."""
        
        processed_range: tuple = self.process_range(cell_range)
        sheet_name: str = cell_range.split("!")[0]
        # Rows start at 0 behind the scenes. +1 match spreadsheet starting at 1.
        next_row: int = (processed_range[2] + 1) + len(rows)

        # Check if you are adding to the end of the sheet.
        # next_row will be greater than current_sheet:next_row if adding to the end.
        if self.current_sheets[sheet_name]["next_row"] < next_row:
            self.current_sheets[sheet_name]["next_row"] = next_row

        self.update_values_requests.append({"range": cell_range, "values": rows})

    
    def change_google_sheet_title_request(self, name: str) -> None:
            "Change the name of the google sheet."

            self.requests.append({"updateSpreadsheetProperties":{"properties":{"title": f"{name}"},"fields": "title"}})


    def change_sheet_name_request(self, old_name: str, new_name: str) -> None:
        """Change the name of the specified sheet."""
        
        self.current_sheets[new_name] = self.current_sheets[old_name]
        del self.current_sheets[old_name]

        self.requests.append({"updateSheetProperties":{
            "properties":{"sheetId": self.current_sheets[new_name]["id"],"title": f"{new_name}"},
            "fields": "title"
        }})


    def add_sheet_request(self, name: str) -> None:
        """Add a new sheet request to the Google Sheet."""
        
        self.current_sheets[name] = {"id": self.sheet_id_runner, "next_row": 0}
        self.sheet_id_runner += 1

        self.requests.append({"addSheet":{"properties":{"sheetId": self.current_sheets[name]["id"], "title": f"{name}"}}})
    

    def resize_request(self, cell_range: str, size: int) -> None:
        """Resize a column or row by specified number of pixels.

        Range formats:
            - Columns: "SheetName!StartColumn:EndColumn "
                - Ex. "Sheet1!A:A" will resize column A.
            - Rows: "SheetName!StartRow:EndRow"
                - Ex. "Sheet1!1:1" will resize row 1.
        """

        processed_range: list = self.process_range(cell_range)
        dimension: str = ""
        start_index: int = 0
        end_index: int = 0

        if processed_range[2] < 0:            
            dimension = "COLUMNS"
            start_index = processed_range[1]
            end_index = processed_range[3]
        else:
            dimension = "ROWS"
            start_index = processed_range[2]
            end_index = processed_range[4]

        format_style = {
            "updateDimensionProperties": {
                "properties": {
                    "pixelSize": size
                },
                "fields": "pixelSize",
                "range": {
                    "sheetId": processed_range[0],
                    "dimension": dimension,
                    "startIndex": start_index,
                    "endIndex": end_index
                }
            }
        }
        self.requests.append(format_style)


    def align_and_wrap_cells_range_request(
        self, 
        cell_range: str, 
        horizontal: str="LEFT", 
        vertical: str="BOTTOM",
        wrapping: str="CLIP"
    ) -> None:
        """Align and wrap cells in the provided range based on alignment provided.

        Alignment:
            - Horizontal: LEFT, CENTER, RIGHT
            - Vertical: TOP, MIDDLE, BOTTOM

        Wrapping: OVERFLOW_CELL, CLIP, WRAP
        """
        processed_range = self.process_range(cell_range)
        format_style = {
            "repeatCell": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": horizontal,
                        "verticalAlignment": vertical,
                        "wrapStrategy": wrapping              
                    }
                },
                "fields": "userEnteredFormat(horizontalAlignment, verticalAlignment, wrapStrategy)"
            }
        }
        self.requests.append(format_style)


    def merge_cells_range_request(self, cell_range: str, merge_type: str="MERGE_ALL") -> None:
        """Merge cells in the provided range based on merge type.
        Merge Types: MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS
        """
        
        processed_range: list = self.process_range(cell_range)
        format_style = {
            "mergeCells": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "mergeType": merge_type
            }
        }
        self.requests.append(format_style)

    
    def format_font_range_request(
        self,
        cell_range: str,
        font_family: str = "Arial",
        font_size: int=12,
        bold: bool = False,
        italic: bool = False,
        strikethrough: bool = False,
        underline: bool = False,
        text_color: tuple = (0, 0, 0)
    ) -> None:
        """Set the font for a range of cells."""

        processed_range = self.process_range(cell_range)
        format_style = {
            "repeatCell": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "foregroundColor": {"red": text_color[0], "green": text_color[1], "blue": text_color[2]},
                            "font_family": font_family,
                            "fontSize": font_size,
                            "bold": bold,
                            "italic": italic,
                            "strikethrough": strikethrough,
                            "underline": underline
                        }
                    }
                },
                "fields": "userEnteredFormat(textFormat)"
            }
        }
        self.requests.append(format_style)


    def fill_range_request(self, cell_range: str, fill_color: tuple=(1, 1, 1)) -> None:
        """Set the background fill for a range of cells.
        Amount of (Red, Green, Blue)
            - Red: 0.0 - 1.0
            - Green: 0.0 - 1.0
            - Blue:  0.0 - 1.0
        """

        processed_range = self.process_range(cell_range)
        format_style = {
            "repeatCell": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {"red": fill_color[0], "green": fill_color[1], "blue": fill_color[2]}                        
                    }
                },
                "fields": "userEnteredFormat(backgroundColor)"
            }
        }
        self.requests.append(format_style)


    def set_outer_border_range_request(self,  cell_range: str, type: str="SOLID", color: tuple=(0, 0, 0)) -> None:
        """Set the outer border for a range of cells.
        Border types:
            - DOTTED
            - DASHED
            - SOLID
            - SOLID_MEDIUM
            - SOLID_THICK
            - NONE
            - DOUBLE
        """

        processed_range = self.process_range(cell_range)
        border_format = {
            "updateBorders": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "top": { "style": type, "color": {"red": color[0], "green": color[1], "blue": color[2]}},
                "bottom": {"style": type, "color": {"red": color[0], "green": color[1], "blue": color[2]}},
                "left": {"style": type, "color": {"red": color[0], "green": color[1], "blue": color[2]}},
                "right": {"style": type,  "color": {"red": color[0], "green": color[1], "blue": color[2]}},
            }
        }
        self.requests.append(border_format)


    def set_bottom_border_range_request(self,  cell_range: str, type: str="SOLID", color: tuple=(0, 0, 0)) -> None:
        """Set the bottom border for a range of cells.
        Border types:
            - DOTTED
            - DASHED
            - SOLID
            - SOLID_MEDIUM
            - SOLID_THICK
            - NONE
            - DOUBLE
        """

        processed_range = self.process_range(cell_range)
        border_format = {
            "updateBorders": {
                "range": {
                    "sheetId": processed_range[0],
                    "startColumnIndex": processed_range[1],
                    "startRowIndex": processed_range[2],
                    "endColumnIndex": processed_range[3],
                    "endRowIndex": processed_range[4]
                },
                "bottom": {"style": type, "color": {"red": color[0], "green": color[1], "blue": color[2]}}
            }
        }
        self.requests.append(border_format)


    def process_range(self, cell_range: str) -> tuple:
        """Process the range into sheet_id and starting and ending cell."""

        range_parts: list = cell_range.split("!")
        sheet_id: int = self.current_sheets[range_parts[0]]["id"]
        start_pair = list()
        end_pair = list()        
        range_pair_values: list = range_parts[1].split(":")

        if len(range_pair_values) > 1:
            start_pair: list = self.process_cell_pair(range_pair_values[0])
            end_pair: list = self.process_cell_pair(range_pair_values[1])
        else:   
            start_pair: list = self.process_cell_pair(range_pair_values[0])
            end_pair: list = [-1, -1]

        return [sheet_id, start_pair[0], start_pair[1]-1, end_pair[0] + 1, end_pair[1]]

        
    def process_cell_pair(self, pair: str) -> list:        
        """Determine the numerical value for the cell locations."""

        base_columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        column_num: int = 0
        row_num: int = -1
        column_char_count: int = 0

        for i, char in enumerate(pair):
            if char in base_columns:
                column_num += base_columns.index(char)
                column_char_count += 1
            else:
                row_num = int(pair[i:])
                break
        
        if column_char_count > 1:
            column_num += 1

        return [column_num, row_num]               


if __name__ == "__main__":
    print("This is a module...")