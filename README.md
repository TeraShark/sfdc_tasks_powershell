# Powershell Scripts: Easy SFDC Task Creation and Integration
## What is this?
A Powerful set of scripts which you can execute whenever you need to create a Task for an Opportunity, List your existing tasks, open them in the browser, sync sharepoint lists from SFDC, etc.

As opposed to navigating through the SFDC Web UI, these scripts allow you to simply enter the numbered options from the script prompts, and you're done. As opposed to spending valuable time navigating through the slow and painful SFDC Web UI, waiting for sections and pages to load, clickety-click-clicking and selecting from a plethora of options, this script gets you through the process quickly and easily :-)

## Script files ##

| Script                      | Function                                                      |
|-----------------------------|---------------------------------------------------------------|
|`sfdc_create_task.ps1`       |Creates tasks in SFDC, prompting for input through the process |
|`sfdc_sync_tasks.ps1`        |Synchronizes tasks created in Sharepoint to SFDC directly |
|`sp_connector.ps1`           |Reads a Sharepoint list, and updates each list item synchronously from SFDC |  

## Installation
### Prerequisites
1. You'll need local administrator rights. This is provided through "BeyondTrust Privilege Management". 
    - If you get a prompt requesting credentials when you select "Run as Administrator", you'll need to submit a request to elevate your "BeyondTrust" permissions through IT. 
2. You will also need the following installed and configured in order to run the SFDC-related scripts in this repo:
    - PowerShell 7.x ([download here, from the official github repo](https://github.com/PowerShell/powershell/releases))
    - Sharepoint PnP PowerShell module
        - Once PowerShell 7 is installed, install the PnP module by running this command in a PowerShell terminal:  
        `Install-Module PnP.PowerShell`
    - SFDC CLI ([download here, from Salesforce](https://developer.salesforce.com/tools/salesforcecli))

### Setup and configuration
1. At the top of each `.ps1` script file (within the first few lines, where variables are instantiated), ensure that your email address for SFDC is listed as the value for `$username`. 
    - NO PASSWORDS!
2. In the 'sfdc_create_task.ps1' script (if you plan on using this script), customize the list of task types (`$taskTypes`) in the task script to match your most commonly used types, ensuring that they match **exactly** what is listed in your SFDC UI instance.
3. Open a Powershell 7 terminal **as an Administrator**, and execute this command:  
`Set-ExecutionPolicy Unrestricted`
4. Open a new, regular (**non-admin**) Powershell 7 terminal, and execute this command:  
`sfdx force:auth:web:login -a dell`
    - This will open a web browser window and prompt you for permission to access the SFDC API, using your SSO credentials - this creates a token for subsequent calls to SFDC. 
    - Follow the sign-in prompts, and once you've signed in successfully, close the browser tab. In the Powershell terminal, you should see a message indicating that your token has been stored - this means you're good to go :-)
    - *No SSO sign-in?* If you see a regular, non-Dell Salesforce sign-in page, there should be a link towrds the bottom of that page for a custom domain. Click the link, and enter `dell` when the prompt comes up, then follow the sign-in process as denoted above.
4. **(*Optional | recommended*)** Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.
    * To create a shortcut, simply right-click an open space on your Desktop, select "New -> Shortcut".
    * For the path, you will need to specify the path to Powershell, followed by the path to the relevant `sfdc_script.ps1` script. 
        * Eg. `"C:\Program Files\PowerShell\7\pwsh.exe" -WorkingDirectory ~ "<path to your>\sfdc_script.ps1">`
    * You can also change the icon. I've included an icon in this repo.
    * In order to create a Taskbar shortcut, first launch your newly created shortcut, right-click on the now-running powershell instance IN THE TASKBAR, and select "Pin to taskbar"

## Usage
Simply launch the shortcut or script using Powershell and follow the prompts. Easy-peazy-lemon-squeezy.