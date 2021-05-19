from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime

def main() -> None:

    # Parameters
    start_date_range = datetime(2020, 1, 1)
    end_date_range = datetime(2021, 2, 21)
    # Use 'All' to search for every training
    training_title = 'Advanced Python Programming - High School CS Professional Development'
    last_cell_of_input_spreadsheet= 'g3443'
    input_spreadsheet_id = '1dZaKc6q3aZM5gPiH1ciKqEeFw6i3DNNvayTagDgUZnY'
    output_spreadsheet_id = '1nYTi7s1VDFXsqupYs7ZPe_9_SrDL2hAAI_i0jnEB5hw'

    # Credentials
    service_account_file= 'need_info.json'
    g_scopes= ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    creds = service_account.Credentials.from_service_account_file(
            service_account_file, scopes=g_scopes)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # Reading values from the spreadsheet
    result_of_read = sheet.values().get(spreadsheetId=input_spreadsheet_id,
                                range=f'Sheet1!a1:{last_cell_of_input_spreadsheet}').execute()
    
    values = result_of_read.get('values', [])

    generate_output_spreadsheet(values, sheet, start_date_range, end_date_range, output_spreadsheet_id, training_title)


def generate_output_spreadsheet(values, sheet, start_date_range, end_date_range, output_spreadsheet_id, training_title) -> None:
    
    values_to_write = [[], [values[0][3], values[0][4], values[0][1], values[0][2], values[0][0]]]
    participants_dict = {}

    # Removing Column headers
    values.pop(0)
    
    participants_dict = process_data_for_records(values, start_date_range, end_date_range, training_title)

    training_numbers = {}

    for record in participants_dict.values():
        values_to_write.append(record)

        if record[3] in training_numbers:
            number = training_numbers[record[3]] + 1
            training_numbers[record[3]] = number
        else:
            training_numbers[record[3]] = 1

    values_to_write.append([])
    total = 0

    for training in training_numbers:
        total += training_numbers[training]
        values_to_write.insert(0, [training, training_numbers[training]])

    values_to_write.insert(0, ['All Trainings', total])
    values_to_write.insert(0, ['Training', 'Total'])

    # Writing Section
    sheet.values().update(spreadsheetId=output_spreadsheet_id, 
                                     range='Sheet1!a1',
                                     valueInputOption='USER_ENTERED',
                                     body={'values':values_to_write}).execute()


def process_data_for_records(values, start_date_range, end_date_range, training_title) -> dict:
    ''' Returns a dictionary that contains an entry for every unique participant/training combo.
    This function will create a new entry for a unique participant/training combo or
    it will append a new timestamp to the end of the participant/training combo record.
    '''
    participants_dict: dict = {}

    for line in values:
        if line[2] == training_title or training_title == 'All':
            timestamp = datetime.strptime(line[0], '%m/%d/%Y %H:%M:%S')

            if timestamp >= start_date_range and timestamp <= end_date_range:
                # Generate a possible key combination that is '<training name><participant last name><participant first name>'
                possible_key = f'{line[2].upper()}{line[4].upper()}{line[3].upper()}'

                # Check if the key is in the dictionary and has the same training.
                if possible_key in participants_dict:
                    # Append the new timestamp, in original form, to the end of the participants current info.
                    participant_info = participants_dict.get(possible_key)
                    participant_info.append(line[0])
                    participants_dict[possible_key] = participant_info
                else:
                    # Create a new key entry with the info: [last_name, first_name, email, training, first_timestamp]
                    participants_dict[possible_key] = [line[4], line[3], line[1], line[2], line[0]]

    return participants_dict

if __name__ == '__main__':
    main()