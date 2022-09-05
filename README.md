# Powershell Script: Easy SFDC Task Creation
## What is it?
A Powerfull script which you can execute whenever you need to create a Task for an Opportunity.

The script relies on the official SFDC CLI tool, downloadable [here](https://developer.salesforce.com/tools/sfdxcli).

As opposed to navigating through the SFDC Web UI, this script allows you to simply enter the Opportunity No (Deal ID), select the task type from a reduced list of tasks, enter a Task description and you're done. As opposed to spending valuable time navigating through the web UI, waiting for sections and pages to load, clickety-click-clicking and selecting from a plethora of options, this script gets you through the process in as little as 15 seconds :-)

## Installation

There are 4 mandatory steps to install this script, and 2 optional steps for a better experience.

1. Download and install the SFDC CLI tool for SFDX by downloading [directly from here](https://developer.salesforce.com/media/salesforce-cli/sfdx/channels/stable/sfdx-x64.exe).
    * ***Note***: You may need Admin rights to install the CLI package.
2. **(*Optional | recommended*)** Download and install v7.2 of Powershell [directly from here](https://github.com/PowerShell/PowerShell/releases/download/v7.2.6/PowerShell-7.2.6-win-x64.msi).
3. Download the files from this repo or clone it.
4. Modify the **"*sfdc_create_task.ps1*"** script to include your corporate email address or username for SFDC.
    * At the top of file, ensure that your email address for SFDC is listed as the value for `$username`.
    * Customize the list of task types (`$taskTypes`) to match your most commonly used types, ensuring that they match what is listed in your SFDC UI instance.
5. Perform an initial request to log in and cache your credentials in the SFDC CLI tool.
    * Open a powershell prompt and type: `sfdx-add-org`
    * This will prompt you to log in through the browser, then enter your Organization alias or Organization ID, and, finally, prompt you to specify whether it is a sandbox or production instance. After this completes, you'll be ready to use this script.
6. **(*Optional | recommended*)** Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.
    * To create a shortcut, simply right-click an open space on your Desktop, select "New -> Shortcut".
    * For the path, you will need to specify the path to Powershell, followed by the path to your `sfdc_create_task.ps1` script. 
        * Eg. `"C:\Program Files\PowerShell\7\pwsh.exe" -WorkingDirectory ~ "<path to your>\sfdc_create_task.ps1">`
    * You can also change the icon. I've included an icon you can use in this repo.
    * In order to create a Taskbar shortcut, first launch your newly created shortcut, right-click on the now-running powershell instance IN THE TASKBAR, and select "Pin to taskbar"

## Usage
Simply launch the shortcut or script using Powershell and follow the prompts. Easy-peasy-lemon-squeezy.