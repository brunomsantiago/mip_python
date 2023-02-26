# Microsoft Information Protection (MIP) in Python

## MIP with openpyxl

Microsoft Information Protection (MIP) labels is an Office 365 feature that allows companies to define custom sensitivity labels (public, internal, ...) to its documents and to require its users to classify each one of them.

It is part of Office 365 but not part of the [Office Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) specification that defines the office documents. According to the MIP [documentation](https://learn.microsoft.com/en-us/information-protection/develop/concept-mip-metadata) the labels are implemented as custom properties, looking like this:

|Key                                                        |Value                               |
|-----------------------------------------------------------|------------------------------------|
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Enabled    |true
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SetDate    |2018-11-08T21:13:16-0800
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Method     |Privileged
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_Name       |Confidential
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_SiteId     |cb46c030-1825-4e81-a295-151c039dbf02
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ContentBits|2
|MSIP_Label_2096f6a2-d2f7-48be-b329-b73aaa526e5d_ActionId   |88124cf5-1340-457d-90e1-0000a9427c9

Where the `2096f6a2-d2f7-48be-b329-b73aaa526e5d` in the name of each custom property is the the label id, which is the main metadata that define the label (public, internal, ...) .

The standard way of using MIP labels is to apply them inside on Office 365 apps like Excel, but there are alternatives like the 'Set-AIPFileLabel' powershell tool provided by Microsoft. This powershell tool can be used from python (see [mip_powershell.py](https://github.com/brunomsantiago/mip_python/blob/main/mip_powershell.py)), but I've found it may make take several seconds to apply the label to each file.

As far as I understand, MIP wasn't designed to be used without an Office 365 account, specially because the ActionId which, according to the documentation, is changed every time a MIP label is applied and may be used for audit purposes. However I've found that if you copy the custom properties from an already labeled file to a new file, it will work without issues.

The code snippet below ([mip_openpyxl.py](https://github.com/brunomsantiago/mip_python/blob/main/mip_openpyxl.py)) demonstrates how to use `openpyxl` to copy the custom properties from a MIP labeled file to a new workbook.

```Python
import openpyxl

# Loading an excel file with sensitivity label (MIP Label)
file_with_mip_label = 'file_with_mip_label.xlsx'
workbook_with_mip_label = openpyxl.load_workbook(file_with_mip_label)

# Creating a new workbook and writing something in ht
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
