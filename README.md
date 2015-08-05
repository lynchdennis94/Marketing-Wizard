# Marketing-Wizard v1.4a
A small desktop application to handle your marketing needs!

# About Marketing Wizard
Marketing Wizard is a small desktop application that is designed to facilitate marketing emails in mass quantities.  By using Microsoft Outlook as an intermediary with any third-party mail service, Marketing Wizard allows a user to define an email template, a list of mail recepients, any attachments for the email, and a personalized subject for the current batch of emails.  The application then throttles down email sending speed to prevent issues with outside mailing services to prevent webmail accounts from becoming locked up.

After the initial commit (which occurred with an already funcitoning program), versioning history will be in the format X.Yz, where
- X = Major version number (when brand new functionality has been added that changes or extends the base program)
- Y = Minor version number (when bug fixes or minor additions to the program have been added)
- z = Current build version letter (incremented with each addition to the program that doesn't consititute a minor version increment)


# Current Features
- Read in email addresses from a Microsoft Excel spreadsheet (Under the heading 'Emails')
- Read in email templates that are defined in a .txt format
- Attach multiple files to group of emails
- Allow for per-batch Subject headings
- Throttling of email speeds to prevent email account lockup
- Integration with any 3rd-party email provider that can interface through Microsoft Outlook
- Dynamic logging of email batch, with errors logged by email address for future reference/removal from spreadsheets

# Known Bugs
- None reported so far

# TODO
- Create multiple sending profiles
- Add email validation to message creation, addresses that fail validation will be logged
- Creation of a preview window that allows user to verify the makeup of the email batch prior to sending out (since we are sending out many emails per batch, it makes best logistical sense to have an in-program preview window as opposed to a Microsoft Outlook preview)
- Integrate Microsoft SQL Server to allow for more reliable data storage and manipulation
- Create module that allows for dynamic email creation within the program
- Create module that allows for accessing database informaiton within the program, as well as adding new recepients and removing old recepients (these will be logged)
- Create a background service that allows for users to read emails into the program (instead of interfacing through the Outlook GUI), and ideally tie these emails to any past marketing emails for statistics and analytics


# Version History
- V1.1a: Added initial project files to source control, created README.md
- V1.1b: Added a preview window that will display the subject and body of the email to be sent out.  Still need to include preview for attachments, as well as provide a way to send emails straight from the preview window if desired (or save any changes made in the preview window to the resulting email).
- V1.2a: Created a utility ExcelRowBinder class to allow for multiple emails from the same excel row to be stored in a dictionary.  Tested multiple-email lines for confirmation they worked correctly
- v1.3a: Added new icon, fixed minor bug with email column having a single email with a comma causing exception.  Also modified the way the row column is retrieved for logging so that the sent column isn't overwritten
- v1.4a: Added a status bar to the program to aid in knowing where in the list the program is
