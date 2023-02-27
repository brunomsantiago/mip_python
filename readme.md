# Microsoft Information Protection (MIP) in Python

## Introduction

Microsoft Information Protection (MIP) labels is an Office 365 feature that allows companies and organizations to define custom sensitivity labels (public, internal, ...) to its documents and to require its users to classify each one of them.

It is part of Office 365 but not part of the [Office Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) specification that defines the office documents. According to the [MIP documentation](https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata) the labels are implemented as [custom properties](https://support.microsoft.com/en-us/office/view-or-change-the-properties-for-an-office-file-21d604c2-481e-4379-8e54-1dd4622c6b75), looking like this:

|Key                                                        |Value                               |
|-----------------------------------------------------------|------------------------------------|
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Enabled    |true
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SetDate    |2018-11-08T21:13:16-0800
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Method     |Privileged
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Name       |Confidential
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SiteId     |cb46c030-1825-4e81-a295-151c039dbf02
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ContentBits|2
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ActionId   |88124cf5-1340-457d-90e1-0000a9427c9

Where the `2096f6a2-d2f7-48be-b329-b73aaa526e5d` in the name of each custom property is the the `label_id`, which is the main metadata that defines the label (public, internal, ...) .

The standard way of using MIP labels is to apply them inside on Office 365 apps like Excel, but there are official alternatives like the [Set-AIPFileLabel](https://learn.microsoft.com/en-us/powershell/module/azureinformationprotection/set-aipfilelabel?view=azureipps) powershell tool provided by Microsoft.


## Method 1: Copying Custom Properties

This method involves copying custom properties from a previously labeled file to a new one. It is simple, fast and does not require external tools nor an Office 365 account. While effective, this method  is non-standard and may have potential drawbacks.
 
I have conducted some tests using Python-generated Excel spreadsheets and have not encountered any issues. The file opens normally, without any warnings, and with the correct sensitivity label applied.

I think the custom properties `MSIP_Label_{label-id}_ActionId` or `MSIP_Label_{label-id}_SetDate` may present potential issues in the future.

According to the documentation `ActionId` is changed every time a MIP label is applied by a standard tool and may be used for audit purposes. It seems to be an UID and it would be easy to generate a new one in Python, but since [the documentation](https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata) mentions audit purposes I suspect it is also sent to the organization's Azure Directory.

The `SetDate` is also changed every time a MIP label is applied and by just copying it from other file it may be inconsistent with the file creation date. It also would be easy to generate a new timestamp with python.

But with no issues so far, I plan to continue using this method of copying all properties, as it is akin to updating the original file with Python.

The code snippet below ([mip_openpyxl.py](https://github.com/brunomsantiago/mip_python/blob/main/mip_openpyxl.py)) demonstrates how to use `openpyxl` to copy the custom properties from a MIP labeled file to a new spreadsheet.

```Python
import openpyxl

# Loading an excel file with sensitivity label (MIP Label)
file_with_mip_label = 'file_with_mip_label.xlsx'
workbook_with_mip_label = openpyxl.load_workbook(file_with_mip_label)

# Creating a new workbook and writing something in it
new_workbook = openpyxl.Workbook()
worksheet = new_workbook.active
cell = worksheet.cell(row=1, column=1)
cell.value = 'Hello World'

# Copying custom properties from one workbook to another
for prop in workbook_with_mip_label.custom_doc_props.props:
    print(f"{prop.name}: {prop.value}")
    new_workbook.custom_doc_props.append(prop)

# Saving the new workbook
new_workbook.save('new_file_with_same_mip_label.xlsx')
```
## Method 2: Call powershell standard tool from  Python

This method involves using a powershell command as a subprocess to apply the sensitivity label. It relies on the [Set-AIPFileLabel](https://learn.microsoft.com/en-us/powershell/module/azureinformationprotection/set-aipfilelabel?view=azureipps) powershell tool, which requires only the `filepath` and `label_id` for the sensitivity label.

Note that some sensitivity labels may have a `MainLabelId` and a `SubLabelId`. In such cases, only the `SubLabelId` should be used as the `label_id` with the `Set-AIPFileLabel` tool, as stated in the [tool's documentation](https://learn.microsoft.com/en-us/powershell/module/azureinformationprotection/set-aipfilelabel?view=azureipps).

The main advantage of this method is it uses an official tool. However you can only use it in a Windows machine within the organization, which may limit some applications. Also I have conducted some tests using this method and found it may make take several seconds to apply the label to each file.

To find the label_id you can inspect the office document custom properties or use other powershell called [Get-AIPFileStatus](https://learn.microsoft.com/en-us/powershell/module/azureinformationprotection/get-aipfilestatus?view=azureipps), which several properties including `MainLabelId` and a `SubLabelId`.

The code snippet below ([mip_powershell.py](https://github.com/brunomsantiago/mip_python/blob/main/mip_powershell.py)) demonstrates two Python functions, one to  find the `label_id` and other apply the sensitivity label to file to.

```Python

import json
import subprocess
import time


def read_label(
    filepath,
    full_result=False,
    powershell=r'C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe',
    stdout_encoding='iso8859-15',
):
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
```