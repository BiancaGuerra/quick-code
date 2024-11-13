import os
import psutil
import subprocess
import time
import win32com.client

''' LOGIN SAP'''

def is_process_running(process_name):
    """Check if a process is running.

    Args:
        process_name (str): The name of the process to check.

    Returns:
        bool: True if the process is running, False otherwise.
    """

    for process in psutil.process_iter(['pid','name']):
        if process.info['name'] == process_name:
            return True
    return False

def get_credentials(credential_file):
    """Obtains the credentials to login.

    Args:
        credential_file (str): A txt file with the structure "username='username' and password='password'".

    Returns:
        Tuple[str, str]: Username and password
    """

    with open(credential_file, 'r') as document:
        contents = document.read().strip().split('\n')
    credentials = {}
    for line in contents:
        key, value = line.split('=')
        credentials[key] = value
    username = credentials.get('username')
    password = credentials.get('password')

    return username, password

def login_sap(credential_file):
    """Open SAP GUI, login (with or without SSO) and open the transaction.

    Returns:
        session (str): The SAP's window we will be using.
    """

    # Open the '.exe' file if it's not already running
    sap_logon_file_path = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe'
    if not is_process_running('saplogon.exe'):  
        subprocess.Popen(sap_logon_file_path)
    else:
        print('SAP Logon is already running')

    # Wait for application to properly load
    for contador in range(10):
        try:
            SapGuiAuto = win32com.client.GetObject('SAPGUI')
        except Exception as e:
            print(f"An error occurred: {e}")
            time.sleep(1)

    username, password = get_credentials(credential_file)

    # Open SAP GUI and login
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection('', True)       # Insert the name of the transaction
    session = connection.Children(0)

    # Fill in credentials if requested
    try:
        session.findById('wnd[0]').maximize
        session.findById('wnd[0]/usr/txtRSYST-MANDT').text = '020'
        session.findById('wnd[0]/usr/txtRSYST-BNAME').text = username
        session.findById('wnd[0]/usr/pwdRSYST-BCODE').text = password
        session.findById('wnd[0]/usr/txtRSYST-LANGU').text = 'EN'
        session.findById('wnd[0]').sendVKey(0)
    except:
        pass

    # Open another window if one is already open
    try:
        session.findById('wnd[1]/usr/radMULTI_LOGON_OPT2').select()
        session.findById('wnd[1]/usr/radMULTI_LOGON_OPT2').setFocus()
        session.findById('wnd[1]/tbar[0]/btn[0]').press()
    except:
        pass

    return session


''' TXT TO CSV'''

def remove_ornaments(file_name, top_lines, bottom_lines):
    """Remove unnecessary lines.

    Args:
        file_name (str): Txt file you want to convert to csv.
        top_lines (int): Number of unnecessary lines at the top of the txt file.
        bottom_lines (int): Number of unnecessary lines at the bottom of the txt file.

    Returns:
        str: String resulted by the join of the corrected data.
    """

    with open(file_name) as file:
        # Read the text and transform each line in one element on the list
        raw_data = file.readlines() 
    
    # Remove the slashes from the beginning
    for i in range(top_lines):
        del raw_data[0]

    # Remove the slashes from the ending
    for i in range(bottom_lines):  
        del raw_data[-1]

    # Remove the slash after the header (when we keep the header)
    if top_lines>3:
        pass
    else:
        del raw_data[1]

    # Transform the list in a string
    return "".join(raw_data)

def pipe_to_virgula(with_pipe):
    """Replace the needed characters.

    Args:
        with_pipe (str): String resulted by the join of the corrected data.

    Returns:
        with_pipe (list): Same data but now in a form of a list with some adjustments.
    """

    with_pipe = with_pipe.replace("|", "", 1)         # Remove the | from the beginning
    with_pipe = with_pipe.replace("\n|", "\n")        # Remove the | from the ending
    with_pipe = with_pipe.split("|")                  # Turn back to a list

    for index, value in enumerate(with_pipe):
        # Remove aditional spaces
        with_pipe[index] = with_pipe[index].strip(" ")

    # Transform the list in a txt file, using ';' as separator
    with_pipe = ";".join(with_pipe)
    
    return with_pipe

def txt_to_csv(extraction, csv, top_lines = 3, bottom_lines = 3):
    """Generate the csv file.

    Args:
        extraction (str): Path to the txt file you want to convert to csv.
        csv (str): Path to the cv file you want to save.
    """
    with_pipe = remove_ornaments(extraction, top_lines, bottom_lines)       # String with no ornaments
    with_comma = pipe_to_virgula(with_pipe)                                 # String with corrected data and separated by commas

    # Write the file
    with open(csv, "w") as file:
        file.write(with_comma)
    
    os.remove(extraction)     # Delete the txt file


''' MAIN '''

def main():
    credential_file = r''       # Insert the path where the credential file is
    extraction = r''      # Insert the path where the extraction were saved (include the name of the file and its extension)
    csv = r''           # Insert the path where you want to save the csv file (also include the name of the file and its extension)

    # Initialize the automation
    beginning_time = time.time()

    # Obtains the credentials and login
    session = login_sap(credential_file)

    # Extraction stages
    ''' Here you can outline the steps needed to extract the required database in SAP.
    You can use SAP Scripting to record every action, which will generate a script.
    You can then copy it here and make any necessary adjustments. '''

    # Convert to csv
    txt_to_csv(extraction, csv)

    # Finaliza a automação
    ending_time = time.time()
    execution_time = ending_time - beginning_time
    minutes = execution_time//60
    seconds = execution_time%60
    print(f"Execution time: {minutes} minutes and {round(seconds,2)} seconds")

if __name__ == "__main__":
    main()