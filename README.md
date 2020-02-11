![alt text](https://github.com/MrJeremyHobbs/CleanSlips-Django/blob/master/cleanslips/static/images/logo_large.png)
# CleanSlips-Django 
A web app for generating interlibrary loan slips in Alma for the CSU system.

# Background
Each campus has individual templates, found in the cleanslips/static/slip_templates folder.

All business logic is in the views.py file located in the cleanslips folder.

The templates are created in Microsoft Word saved as docx files. They make use of document columns, instead of table columns.

The output will be in the OpenOffice docx format.

# How Does it Work?
Slips are generated by performing a mailmerge against the LendingRequest.xls input file.

CleanSlips uses the [MailMerge](https://pypi.org/project/docx-mailmerge/) python module.

A good tutorial on this module can be found [here](https://pbpython.com/python-word-template.html).

# Other Versions
A legacy desktop app is available [here](https://github.com/MrJeremyHobbs/CleanSlips).