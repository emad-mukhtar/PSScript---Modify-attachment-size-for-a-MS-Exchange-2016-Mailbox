# README for Exchange Attachment Size Modifier Script
## Introduction
This GitHub repository contains a script to modify inbound and outbound email attachment size for a single user in MS Exchange 2016 on-premises. The script requires the user to enter the Exchange server name, username, password, AD domain name, and email domain name. Once the information has been entered, the script will prompt the user for confirmation and then modify the MaxReceiveSize and MaxSendSize of the specified user. The script includes fail-safe checks to ensure the new attachment size does not exceed the user's mailbox size or quotas.

## Prerequisites
The code is written in PowerShell and uses the ExchangeOnlineManagement module. To use this script, you must have:

* A computer with Windows OS and PowerShell installed
* A valid Exchange Online or Exchange 2016 on-premises account
* Suffecient MS Exchange user permissions
* Access to the Exchange Management Shell

## Usage
You can either run the script from your computer using PowerShell to connect to MS Exchange 2016 on-premises or log in directly to MS Exchange 2016 and run the script from the Exchange PowerShell.

* Download the script from this GitHub repository.
* Open PowerShell on your computer or log in to Exchange 2016 on-premises.
* Enter the required information when prompted by the script, including the Exchange server name, username, password, AD domain name, and email domain name.
* Review the changes that will be made and confirm the prompt.
* The script will run and modify the MaxReceiveSize and MaxSendSize for the specified user.
## Author
The script was created by Emad Mukhtar and is available under the MIT license.

## Note
If you have any questions or feedback, please visit my GitHub at https://github.com/emad-mukhtar.

Please use this script at your own risk. I'm not responsible for any damage or loss caused by the use of this script. It is highly recommended to backup your Exchange data before running this script.
