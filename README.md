# Powershell Script: Easy SFDC Task Creation
## What is it?
A Powerfull script which you can execute whenever you need to create a Task for an Opportunity.

The script relies on the official SFDC CLI tool, downloadable [here](https://developer.salesforce.com/tools/sfdxcli).

As opposed to navigating through the SFDC Web UI, this script allows you to simply enter the Oppoirtunity No, select the task type from a reduced list of tasks, enter a Task description and you're done. As opposed to spending valuable time navigating through the web UI, waiting for esctions to load, clickety-click-clicking and selecting from a plethora of options, this script gets you through the process in as little as 15 seconds :-)

## Installation

There are 5 steps required to install this script.

1. Download and install the SFDC CLI tool for SFDX by downloading [directly from here](https://developer.salesforce.com/media/salesforce-cli/sfdx/channels/stable/sfdx-x64.exe).
2. (*Optional but recommended*) Download and install v7.2 of Powershell [directly from here](https://github.com/PowerShell/PowerShell/releases/download/v7.2.6/PowerShell-7.2.6-win-x64.msi).
3. Download the files from this repo or clone it.
4. Modify the "*sfdc_create_task.ps1*" script to include your corporate email address or username for SFDC.
5. Perform an initial request to log in and cache your credentials in the SFDC CLI tool.
6. (*Optional*) Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.

## Usage
