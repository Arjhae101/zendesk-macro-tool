# Zend Macro Tool

## Setup
1. Open the [Google Spreadsheet](https://docs.google.com/spreadsheets/d/1TaKFqCdorojgzJFctZA_Kf5FkbFgRy_H7UuwUWkhGA4) and make a copy.
1. Update the code in the "Script Editor" with the latest version of the [zendesk-macro-tool.gs](google-script/zendesk-macro-tool.gs) file from github
1. Create a token for your account in Zendesk
1. Enter your credentials, your username is your Zendesk account with `/token` at the end

## Usage
### Hints & Limitations
"Import Macros" will import all data again from Zendesk and overwrite all your changes!
The grey fields are currently not synced back and should not be edit
Changes should be synced back to Zendesk in smaller chunks (10-20 rows), the API limits the number of calls
Only the fields Subject and Comment will be synced back, everything eles wont be synced, but also not overwritten in Zendesk

### Editing Process
Edit "Subject" and "Comment" content
Set the "Status" to "edited" after you finished

### Syncing back to Zendesk
The "Sync back to Zendesk" function syncs all rows which have the status "edited" back to Zendesk
Afterwards the rows will have the status "synced-to-zendesk"