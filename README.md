# Powershell Script: Easy SFDC Task Creation
## What is it?
A Powerful set of scripts which you can execute whenever you need to create a Task for an Opportunity, List your existing tasks, open them in the browser, sync sharepoint lists from SFDC, etc.

As opposed to navigating through the SFDC Web UI, these scripts allow you to simply enter the numbered options from the script prompts, and you're done. As opposed to spending valuable time navigating through the slow and painful SFDC Web UI, waiting for sections and pages to load, clickety-click-clicking and selecting from a plethora of options, this script gets you through the process quickly and easily :-)


## Installation
### The easier way
Run the `setup.ps1` script AS AN ADMINISTRATOR and it will prompt you through the process.
It will also download whatever you need to complete the installation.
Once setup is complete, don't forget to modify the other (`*.ps1`) scripts to include your corporate email address or username for SFDC, and perhaps customize your task type list, as listed below:

1. At the top of each `.ps1` script file (within the first few lines, where variables are instantiated), ensure that your email address for SFDC is listed as the value for `$username`. 
    - NO PASSWORDS!
2. Customize the list of task types (`$taskTypes`) to match your most commonly used types, ensuring that they match what is listed in your SFDC UI instance.
3. Open a Powershell prompt window. Type in `sfdx force:auth:web:login` and hit enter. 
    - This will open a web browser window and prompt you for permission to access the SFDC API, using your SSO credentials - this creates a token for subsequent calls to SFDC. 
    - Follow the prompts and log in on that window, then close the browser and the Powershell window.
4. **(*Optional | recommended*)** Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.
    * To create a shortcut, simply right-click an open space on your Desktop, select "New -> Shortcut".
    * For the path, you will need to specify the path to Powershell, followed by the path to the relevant `sfdc_script.ps1` script. 
        * Eg. `"C:\Program Files\PowerShell\7\pwsh.exe" -WorkingDirectory ~ "<path to your>\sfdc_script.ps1">`
    * You can also change the icon. I've included an icon you can use in this repo.
    * In order to create a Taskbar shortcut, first launch your newly created shortcut, right-click on the now-running powershell instance IN THE TASKBAR, and select "Pin to taskbar"

### The manual way, in case the setup script above isn't working for you...
Follow these steps to install this script, and 2 optional steps for a better experience:

1. Download and install the SFDC CLI tool for SFDX by downloading [directly from here](https://developer.salesforce.com/media/salesforce-cli/sfdx/channels/stable/sfdx-x64.exe).
    * ***Note***: You may need Admin rights to install the CLI package.
2. **(*Optional | recommended*)** Download and install the latest version of Powershell [from here](https://aka.ms/powershell-release?tag=stable).
3. Now, for the scripts, download the files from this repo or clone it.
4. At the top of each `.ps1` script file (within the first few lines, where variables are instantiated), ensure that your email address for SFDC is listed as the value for `$username`. 
    - NO PASSWORDS!
    - If required, customize the list of task types (`$taskTypes`) to match your most commonly used SFDC tasks, ensuring that they match what is listed in the SFDC UI instance.
5. Perform an initial request to log in and cache your auth token in the SFDC CLI tool.
    * Open a powershell prompt and type: `sfdx force:auth:web:login -a <org>` where `<org>` is your organization alias (eg. microsoft).
    * This will prompt you to log in through the browser, then enter your Organization alias or Organization ID, and, finally, prompt you to specify whether it is a sandbox or production instance. After this completes, you'll be ready to use this script.
6. Set the local Powershell Execution-Policy in order to run sfdx commands
    * ***Note***: You will need **Administrator** rights for this step.
    * Open a Powershell prompt as an **Administrator** and type: `Set-ExecutionPolicy Unrestricted` and then accept the prompt. *If you see errors with this, then try* `Set-ExecutionPolicy RemoteSigned`.
7. **(*Optional | recommended*)** Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.
    * To create a shortcut, simply right-click an open space on your Desktop, select "New -> Shortcut".
    * For the path, you will need to specify the path to Powershell, followed by the path to the relevant `sfdc_script.ps1` script. 
        * Eg. `"C:\Program Files\PowerShell\7\pwsh.exe" -WorkingDirectory ~ "<path to your>\sfdc_script.ps1">`
    * You can also change the icon. I've included an icon you can use in this repo.
    * In order to create a Taskbar shortcut, first launch your newly created shortcut, right-click on the now-running powershell instance IN THE TASKBAR, and select "Pin to taskbar"


## Usage
Simply launch the shortcut or script using Powershell and follow the prompts. Easy-peazy-lemon-squeezy.