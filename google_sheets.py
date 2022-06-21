from googleapiclient.discovery import build
from google.oauth2 import service_account
from datetime import datetime
from collections import OrderedDict
from operator import itemgetter


def main() -> None:

    # Parameters
    start_date_range: object = datetime(2022, 6, 12)
    end_date_range: object = datetime(2022, 6, 25)

    days = [
        datetime(2022, 6, 13),
        datetime(2022, 6, 14),
        datetime(2022, 6, 15),
        datetime(2022, 6, 16),
        datetime(2022, 6, 17),
        datetime(2022, 6, 20),
        datetime(2022, 6, 21),
        datetime(2022, 6, 22),
        datetime(2022, 6, 23),
        datetime(2022, 6, 24)
    ]

    # Use 'All' for training_title to search for every training.
    training_title: str = 'Arkansas Computer Science and Computing Educator Academy'

    sheet_name: str = "Form Responses 1"
    input_spreadsheet_id: str = "10Gl4V7Xa1sx2jBDo0q5av9mow__XvAvxZ_LmhKwndzQ"
    output_spreadsheet_id: str = "1RsqzQ_Jx2q1qYreqKV4UqRM4CUGq4tEDSrGHM5QSinI"

    # Credentials
    service_account_file: str = "google_info.json"
    g_scopes: list = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = None
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=g_scopes
    )

    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    # Reading values from the spreadsheet
    result_of_read = (
        sheet.values()
        .get(
            spreadsheetId=input_spreadsheet_id,
            range=f"{sheet_name}"
        )
        .execute()
    )

    values = result_of_read.get("values", [])

    generate_output_spreadsheet(
        values,
        sheet,
        start_date_range,
        end_date_range,
        output_spreadsheet_id,
        training_title,
        days
    )


def generate_output_spreadsheet(
    values,
    sheet,
    start_date_range,
    end_date_range,
    output_spreadsheet_id,
    training_title,
    days
) -> None:

    header_line = [
            values[0][4],
            values[0][3],
            values[0][1],
            values[0][2]
    ]

    all_attendance_times = []
    for day in days:
        for _ in range(3):
            header_line.append(day.strftime("%m/%d/%Y"))
            all_attendance_times.append(day)

    values_to_write = [
        [],
        header_line,
    ]
    participants_dict = {}

    # Removing Column headers
    values.pop(0)

    participants_dict = process_data_for_records(
        values, start_date_range, end_date_range, training_title
    )

    formatted_participant_dict = {}

    for participant in participants_dict:
        formatted_info = participants_dict[participant][:4]
        participant_times = participants_dict[participant][4:]
        for time in all_attendance_times:
            if len(participant_times) != 0:
                if participant_times[0].date() == time.date():
                    stamp = participant_times.pop(0).strftime("%m/%d/%Y %H:%M:%S")
                    formatted_info.append(stamp)
                elif participant_times[0].date() < time.date():
                    pass
                else:
                    formatted_info.append("------")
            else:
                formatted_info.append("")
        formatted_participant_dict[participant] = formatted_info


    training_numbers = {}

    for record in formatted_participant_dict.values():

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

    values_to_write.insert(0, ["All Trainings", total])
    values_to_write.insert(0, ["Training", "Total"])

    # Clear old data
    sheet.values().clear(
        spreadsheetId=output_spreadsheet_id,
        range=f"Sheet1"
    ).execute()

    # Writing Section
    sheet.values().update(
        spreadsheetId=output_spreadsheet_id,
        range="Sheet1!a1",
        valueInputOption="USER_ENTERED",
        body={"values": values_to_write},
    ).execute()


def process_data_for_records(
    values, start_date_range, end_date_range, training_title
) -> dict:
    """Returns a dictionary that contains an entry for every unique participant/training combo.
    This function will create a new entry for a unique participant/training combo or
    it will append a new timestamp to the end of the participant/training combo record.
    """

    participants_dict: dict = {}

    for line in values:
        if line[2] == training_title or training_title == "All":
            timestamp = datetime.strptime(line[0], "%m/%d/%Y %H:%M:%S")

            if timestamp >= start_date_range and timestamp <= end_date_range:
                # Generate a possible key combination that is '<training name><participant last name><participant first name>'
                possible_key = f'{line[2].strip().upper()}{line[4].strip().upper()}{line[3].strip().upper()}'
                # Generate a possible key combination that is '<training name><participant email>'
                # possible_key = f"{line[2].strip().upper()}{line[1].strip().upper()}"

                # Check if the key is in the dictionary and has the same training.
                if possible_key not in participants_dict:
                    # Create a new key entry with the info: [last_name, first_name, email, training, first_timestamp]
                    
                    participants_dict[possible_key] = [
                        line[4],
                        line[3],
                        line[1],
                        line[2],
                        timestamp,
                    ]
                else:
                    # Append the new timestamp, in original form, to the end of the participants current info.
                    participant_info = participants_dict.get(possible_key)
                    participant_info.append(timestamp)
                    participants_dict[possible_key] = participant_info

    sorted_participants_dict = OrderedDict(
        sorted(participants_dict.items(), key=itemgetter(1))
    )

    return sorted_participants_dict


if __name__ == "__main__":
    main()
