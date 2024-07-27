import pandas as pd
import requests
import json
import re
from pymarc import MARCReader, Record, Field, Subfield
from pymarc import MARCWriter
from requests.auth import HTTPBasicAuth
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import BackendApplicationClient
import yaml
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox

def run_script():
    re_pattern = re.compile(r'\d+')

    # Load configuration from config.yml
    with open("config.yml", 'r') as stream:
        config = yaml.safe_load(stream)

    # Choose the MARC file using a file dialog
    marc_path = filedialog.askopenfilename(title="Select MARC File", filetypes=[("MARC files", "*.mrc;*.mrk")])
    
    # REGEX to save the name of the selected file to a variable called filename_match
    filename_match = re.search(r'[^\\/]+$', marc_path)

    # Checking that the REGEX found a match. When it does, it groups the match object into a variable called filename which can be used to name the updated file later in the script. 
    if filename_match:
        filename = filename_match.group()
        print(filename)
    else:
        print("No filename found.")
        
    #If a file to process is not selected, the script quits. 
    if not marc_path:
        return

    # OCLC service URL statement
    serviceURL = config.get('metadata_service_url')

    # OCLC auth token information
    scope = ['wcapi:view_bib']
    auth = HTTPBasicAuth(config.get('key'), config.get('secret'))
    client = BackendApplicationClient(client_id=config.get('key'), scope=scope)
    wskey = OAuth2Session(client=client)
    
    # Open the MARC record that is being processed.
    with open(marc_path, 'rb') as file:
        reader = MARCReader(file, to_unicode=True, force_utf8=True)

        # Create a new MARC file to store all records
        updated_file_path = f"updated_{filename}"
        with open(updated_file_path, 'ab') as updated_file:
            writer = MARCWriter(updated_file)
            
            # Loop through the MARC record and extract the 020$a from each record. Clean up the extra language stored in the 020 and then save the clean ISBN to a variable called clean_isbn for use. 
            for record in reader:
                try:
                    isbn_field = record['020']

                    isbn_str = str(isbn_field)
                    full_isbn = ''.join(re.findall(re_pattern, isbn_str))
                    clean_isbn = full_isbn[3:]
                    print(clean_isbn)
                    
                    # Get the auth token.
                    try:
                        token = wskey.fetch_token(token_url=config.get('token_url'), auth=auth)
                        
                        # Use the auth token to call the service and pass it the clean_isbn.
                        try:
                            r = wskey.get(serviceURL + f"q={clean_isbn}", headers={"Accept": 'application/json;content="application/json"'})
                            print(r)
                            status = r.raise_for_status()
                            parsed = json.loads(r.content)
                            
                            # Check to see if the service returned a matching record. If not, skip the record and continue the loop. 
                            if parsed.get('numberOfRecords', 1) == 0:
                                print(f"Skipping query {clean_isbn} as numberOfRecords is 0")
                                continue
                                
                            # Where the service returns a matching record, parse the returned JSON to extract the Work ID associated with the record and save it to a variable called work_id
                            oclc_number = parsed['bibRecords'][0]['identifier']['oclcNumber']
                            print(oclc_number)

                            # Add the oclc_number to the 941 field in the MARC record
                            new_field = Field(tag='941', indicators=[' ', ' '], subfields=[Subfield(code='f', value=oclc_number)])
                            record.add_field(new_field)

                        except KeyError as e:
                            print('skipped')
                            # Handle the KeyError (field or subfield not found in the record) by skipping the record and continuing the loop.
                            continue

                    except requests.RequestException as e:
                        print(f"ISBN {clean_isbn} - Request Exception: {e}")

                except KeyError as e:
                    print(f'Error: {e}')
                    # Handle the KeyError (field not found in the record) by skipping the record and continuing the loop.
                    continue

                # Write the updated record to the new MARC file
                writer.write(record)

# Create the GUI
root = tk.Tk()
root.title("Work ID Lookup Tool (MARC)")

# Add a label with an explanation of the script
explanation_text = "This script queries the OCLC metadata search API service using ISBNs from a user-selected MARC file. The 'Work ID' is then extracted from the query results and saved in the an updated version of the user-selected MARC record in the 941$f."
label = tk.Label(root, text=explanation_text, wraplength=300, justify="left")
label.pack(pady=10)

# Add a button to run the script
run_button = tk.Button(root, text="Find Work ID", command=run_script)
run_button.pack(pady=10)

# Run the GUI
root.mainloop()
