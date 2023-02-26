
import json
import subprocess
import time


def read_label(
    filepath,
    full_result=False,
    powershell=r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe',
    stdout_encoding='iso8859-15',
):
    """
    Read sensitivity label from a Microsoft document
    This function uses a powershell command as subprocess to read the label_id
    of a microsoft document previously classified with the sensitivity label.
    This label_id can be used to apply the same sensitivity label to other
    documents.
    It relies on the 'Get-AIPFileStatus' powershell tool. To understand it
    better try running the command directly in powershell or look for the
    official Microsoft documentation.
    By default this function only returns the label_id, but if you want to see
    the full result from 'Get-AIPFileStatus' use full_result=True.
    """
    # The command to call in powershell. It includes the powershell tool
    # 'ConvertTo-Json' to make it easier to process the results in Python,
    # specially when the file path is too long, which may break lines.
    command = f"Get-AIPFileStatus -path '{filepath}' | ConvertTo-Json"
    # Executing it
    result = subprocess.Popen([powershell, command], stdout=subprocess.PIPE)
    result_lines = result.stdout.readlines()
    # Processing the results and saving to a dictionary
    clean_lines = [
        line.decode(stdout_encoding).rstrip('\r\n') for line in result_lines
    ]
    json_string = '\n'.join(clean_lines)
    result_dict = json.loads(json_string)
    # If selected, return the full results dictionary
    if full_result:
        return result_dict
    # If not returns only the label_id of interest to apply to other document
    # Per Microsoft documentation if a sensitivity label has both a
    # 'MainLabelId' and a 'SubLabelId', only the 'SubLabelId' should be used
    # with 'Set-AIPFileLabel' tool to to set the label in a new document.
    label_id = (
        result_dict['SubLabelId']
        if result_dict['SubLabelId']
        else result_dict['MainLabelId']
    )
    return label_id


def apply_label(
    filepath,
    label_id,
    powershell=r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe',
    stdout_encoding='iso8859-15',
):
    """
    Apply sensitivity label to a Microsoft document
    This function uses a powershell command as subprocess to apply it.
    It relies on the 'Set-AIPFileLabel' powershell tool. To understand it
    better try running the command directly in powershell or look for the
    official Microsoft documentation.
    Per Microsoft documentation if a sensitivity label has both a
    'MainLabelId' and a 'SubLabelId', only the 'SubLabelId' should be used
    with 'Set-AIPFileLabel' tool to to set the label in a new document.
    The function returns the elapsed time to apply the label.
    """
    start = time.time()
    # The command to call in powershell
    command = f"(Set-AIPFileLabel -path '{filepath}' -LabelId '{label_id}').Status.ToString()"
    # Executing it
    result = subprocess.Popen([powershell, command], stdout=subprocess.PIPE)
    result_message = (
        result.stdout.readline().decode(stdout_encoding).rstrip('\r\n')
    )
    # If the command is not successful, raises an exception and display the
    #  message from 'Set-AIPFileLabel' tool
    if result_message != 'Success':
        raise Exception(result_message)
    end = time.time()
    return end - start