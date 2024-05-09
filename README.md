# M365-User-Properties-Report-and-Update

See User guide (pdf) for more information: https://github.com/ITAutomator/M365-User-Properties-Report-and-Update/blob/main/M365%20User%20Properties%20Report%20and%20Update%20Readme.pdf

To Download: Click the green Code button (above) and click Download Zip 

Update M365 user properties (Entra properties) in bulk via csv file.

![image](https://github.com/ITAutomator/M365-User-Properties-Report-and-Update/assets/135157036/b0d4e774-e69f-48f1-adca-81b6957d2412)

How it works
Use this code in 2 phases to create a CSV report of the editable properties of your users in Entra.

Phase 1: Report
Run the M365UserPropertiesReport.ps1 (or .cmd) and enter your admin credentials.
This will output a CSV file containing your users.
Note: Only Enabled accounts are reported.  Only members are reported (vs guests).

Phase 2: Edit

![image](https://github.com/ITAutomator/M365-User-Properties-Report-and-Update/assets/135157036/23ddc22b-469d-44ed-8f97-f435e5909e93)

Make a copy of the CSV file.  Put the UserPrincipalName as the first (required) column.
Then, for any properties you would like to update, make any changes you need.
Delete any columns that you don’t want to update.
Start conservatively, with one or two columns to update.

If you adjust the value, the program will adjust the property in Entra
If you leave the value, or change it to blank, the contents will not be changed.
If you enter the keyword ‘<clear>’ (without the quotes) the property will be cleared.

Run the M365UserPropertiesUpdate.ps1 (or .cmd) to make your updates.
The program allows you to step through each user if you want to go slowly.

Properties
Here are the properties (so far) that seem to be editable by this method.

DisplayName, mail,  BusinessPhones, city, country, department, GivenName, JobTitle,  MobilePhone, OfficeLocation, postalcode, state, streetAddress,

Note: BusinessPhones is a multi-valued property.  This code only handles one value for this property.
