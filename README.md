# ManageEngineAPI

This program is a powerfull rest API script developed by Jax Sutton and Connor Pink.

LocalToRegional.ps1 file is a powershell rest API script for operating on Manage Engine to take local tickets and create regional tickets and then update the local ticket to have the information of the regional ticket. This script is hosted on the server and triggered automatically to run comepletely seamlessly. 
The script can handle attached files for tickets such as images, recordings, and documents by downloading the file locally, reencoding the file locally and uploading it back to the regional page. 
Alot of data cleaning was needed for the description due to the way that the systems format the text using html tags and outlook emails sent from iphones being turned into tickets have strange formatting.
In the case of an error happening the script will add a note to the ticket to let the user know that the ticket failed and needs manual creation.

This script was successfully used over 500 times in a 1 month period and saved alot of time.



