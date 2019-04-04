# usershare-report
This PowerShell script will target domain joined computers and measure the size of user profiles on them. It will spit the data to a local Access DB for cleanup and reporting. It is meant to assist in forecasting storage needs if local user profiles are moved to shared storage.

You are meant to run GetUserFilesSize.ps1 only you will need to edit the script and replace the Get-AdComputer Filters with ones which will target your org's computers appropriately. In my case, all computers had either "LT" for laptop or "PC" for desktops, so I filtered for those because I didn't want to pull anything from domain controllers or servers.
