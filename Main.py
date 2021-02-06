from Config import input_data
from Applicant_Builder import ApplicantBuilder
from time import sleep

def main():   
    process_1 = ApplicantBuilder(input_data["sheet_scopes"],
                    input_data["service_file"],
                    input_data["drive_scopes"],
                    input_data["doc_scopes"])
    # Read desired spreadsheet
    process_1.read_sheet(input_data["sheet_read_ID"],
                input_data["sheet_read_range"])
    # Generate dictionary from read data
    process_1.create_dictionary()
    # Create folder and grant permissions
    process_1.create_folder(input_data["folder_name"])
    for email in input_data["user_emails"]:
        process_1.create_permissions(email)
    # Create a document for each index in the dictionary
    index = 0
    for key, value in process_1.get_dictionary().items():
        process_1.create_doc(key)
        docID = process_1.get_doc_ID()[index]
        index += 1
        # Write to document with fields extracted from the Sheets
        process_1.write_doc(key ,
                        process_1.get_dictionary()[key],
                        docID)
        # Print progress
        print("Document Created for ", key, "\n", 
        int(index / len(process_1.get_dictionary() ) * 100 ), "% complete")
        # Sleep due to Google API request quotas
        sleep(0.75)
    return

if __name__ == "__main__":
    main()

