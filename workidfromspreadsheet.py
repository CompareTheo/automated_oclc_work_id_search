import pandas as pd
import requests
import json
import sys, os
from requests.auth import HTTPBasicAuth
from requests_oauthlib import OAuth2Session
from oauthlib.oauth2 import BackendApplicationClient
import yaml
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
    # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def run_script():
    # Load configuration from config.yml
    with open(resource_path("config.yml", 'r')) as stream:
        config = yaml.safe_load(stream)

    # Choose the spreadsheet using a file dialog
    spreadsheet_path = filedialog.askopenfilename(title="Select Spreadsheet", filetypes=[("Excel files", "*.xlsx;*.xls")])
    
    if not spreadsheet_path:
        return

    # Load query strings from the selected spreadsheet
    queries_df = pd.read_excel(spreadsheet_path)

    # Add a new column 'Work ID' to the DataFrame
    queries_df['Work ID'] = ''
    
    # OCLC Service URL Statement
    serviceURL = config.get('metadata_service_url') 
    
    # get a token
    scope = ['wcapi:view_bib']
    auth = HTTPBasicAuth(config.get('key'), config.get('secret'))
    client = BackendApplicationClient(client_id=config.get('key'), scope=scope)
    wskey = OAuth2Session(client=client)

    progress_bar["maximum"] = len(queries_df)

    try:
        token = wskey.fetch_token(token_url=config.get('token_url'), auth=auth)

        for index, row in queries_df.iterrows():
            query_string = row['EAN-13']  # Replace 'QueryColumn' with the actual column name in your spreadsheet
            try:
                r = wskey.get(serviceURL + f"q={query_string}", headers={"Accept": 'application/json;content="application/json"'})
                # print(r)
                status = r.raise_for_status()
                parsed = json.loads(r.content)
                
                if parsed.get('numberOfRecords', 1) == 0:
                    print(f"Skipping query {query_string} as numberOfRecords is 0")
                    continue
                
                work_id = parsed['bibRecords'][0]['work']['id']
                # print(work_id)

                # Save the 'Work ID' to the new column
                queries_df.at[index, 'Work ID'] = work_id

                # Update progress bar
                progress_bar["value"] = index + 1
                root.update_idletasks()

            except requests.exceptions.HTTPError as err:
                print(err)
    except BaseException as err:
        print(err)

    # Save the updated DataFrame back to the spreadsheet
    queries_df.to_excel(spreadsheet_path, index=False)

    # Reset progress bar
    progress_bar["value"] = 0

    # Show a message box when the script is finished
    messagebox.showinfo("Script Finished", "The script has finished processing!")

# Create the GUI
root = tk.Tk()
root.title("Work ID Lookup Tool")
icon = tk.PhotoImage(file=(resource_path("cloud.png")))
root.iconphoto(True, icon)

# Add a label with an explanation of the script
explanation_text = "This script queries the OCLC metadata search API service using ISBNs from a user-selected spreadsheet.\n\n"\
                   "The ISBNs in the spreadsheet must belong to a column labeled EAN-13, as this is the standard format from the LMU bookstore.\n\n"\
                   "The 'Work ID' is then extracted from the query results and saved in the user-selected spreadsheet under a column named Work ID."
label = tk.Label(root, text=explanation_text, wraplength=300, justify="left")
label.pack(pady=10)

# Add a button to run the script
run_button = tk.Button(root, text="Find Work ID", command=run_script)
run_button.pack(pady=10)

# Add a progress bar
progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Run the GUI
root.mainloop()