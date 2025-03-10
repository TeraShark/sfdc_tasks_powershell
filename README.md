# Powershell Scripts: Easy SFDC Task Creation and Integration
## What is this?
A Powerful set of scripts which you can execute whenever you need to create a Task for an Opportunity, List your existing tasks, open them in the browser, sync sharepoint lists from SFDC, etc.

As opposed to spending valuable time navigating through the slow and painful SFDC Web UI, waiting for sections and pages to load, clickety-click-clicking and selecting from a plethora of options, this script gets you through the process quickly and easily :-)

## Dependencies
The scripts in this repo have the following dependencies:
1. PowerShell 7.x
    - With Unrestricted or RemoteSigned execution policy
2. PnP.PowerShell
    - Sharepoint module for PowerShell 7.x
    - The `sp_connector.ps1` script will automatically attempt to install this module dynamically, if not present
3. SFDC CLI (MSI package) from SalesForce.com

## Script files ##

| Script                      | Function                                                      |
|-----------------------------|---------------------------------------------------------------|
|`sfdc_create_task.ps1`       |To quickly and easily create individual Tasks against Opportunities in SFDC. Prompts for input through the process |
|`sp_connector.ps1`           |Reads the Sharepoint Opp Tracker list, and updates those items from SFDC |  
|`sfdc_sync_tasks.ps1`        |Synchronizes tasks created in the Sharepoint Tracker to SFDC directly. The script `sp_connector.ps1` also provides an option to execute this script automatically after syncing Deals |

## Config files ##

When you first execute `sp_connector.ps1`, you will be prompted for your Sharepoint username (usually 'Last Name, First Name'), and your SFDC username (usually your Dell email address, e.g. 'j_soap@dell.com'). These values will be stored (without passwords) in config files named `sharepoint.cfg` and `user.cfg`, respectively. These config files are auto-generated if they don't exist.
There should be no need to modify these auto-generated config files, unless you accidentally mistyped the values when prompted on first run, possibly causing authentication issues.

## Passwords / Authentication ##
**No passwords** are stored in the files in this repo, nor in the config files generated on first execution of scripts. Authentication against SFDC is provided natively by the SFDC CLI, stored as tokens in the CLI subsystem. The scripts in this repo do not ever access such data directly - the official SFDC CLI handles that seamlessly after first login.
For Sharepoint authentication, in-context Windows credentials are leveraged, based on the user executing the scripts.
You will not be prompted for passwords directly in the scripts (only you User ID(s) / User Names(s) for the relevant integrations). By no means, should you ever be storing passwords anywhere in scripts or config files, since this is not only stupid, but will likely be in violation of any company policy. 
You have been warned.

## Installation

### Prerequisites
1. You will, most likely, need local administrator rights to install the required packages and to set the Execution Policy for PowerShell.
2. You will need the following dependencies installed and configured in order to run the scripts in this repo:
    - **Git** for Windows 
        - Install through "*Company Portal*" by searching for "*Git*"
        - You don't necessarily need Git if you choose to download the files from this repo directly, however, if you want the ability to receive updates / refinements to the scripts as I continue to maintain them, I'd suggest you install Git and use it to retrieve the scripts in this repo using '`git clone ...`', and get updates in future, by using '`git pull`'
    - **PowerShell 7.x**
        - Install through "*Company Portal*" by searching for "*Powershell 7*"
   - **SFDC CLI** 
        - [Download here, from Salesforce.com](https://developer.salesforce.com/tools/salesforcecli)
        - Run the installer and follow the prompts
        - *Trouble* running the installer? *Workaround*:
            - Install the latest version of **Node.js** through "*Company Portal*"
            - Once installed, open a Terminal window, and verify the installation by typing the below, which should display the version number and not an error:<br />
            `node --version`
            - Now, type the following into the Terminal window to install the SalesForce CLI: <br />
            `npm install @salesforce/cli --global`

### Setup and configuration
1. Open a Powershell 7 terminal **as an Administrator**, and execute this command:  
`Set-ExecutionPolicy Unrestricted`
2. Close the Admin terminal. Open a new, regular (**non-admin**) Powershell 7 terminal, and execute this command:  
`sfdx force:auth:web:login -a dell`
    - This will open a web browser window and prompt you for permission to access the SFDC API, using your SSO credentials - this creates a token for subsequent calls to SFDC. 
    - Follow the sign-in prompts (you may need to sign in with SSO using the link at the bottom of the page), and once you've signed in successfully, close the browser tab. In the Powershell terminal, you should see a message indicating that your token has been stored - this means you're good to go :-)
    - *No SSO sign-in?* If you see a regular, non-Dell/non-SSO Salesforce sign-in page prompting for a username and password, look for a link towards the bottom of the login page containing "custom domain". Click the link, and enter `dell` for the domain when the prompt comes up. This should launch the SSO process.
3. If you plan on using the standalone script 'sfdc_create_task.ps1' to create SFDC tasks, customize the list of task types (`$taskTypes`) in the task script to match your most commonly used types, ensuring that they match **exactly** what is listed in your SFDC UI instance.
4. (***Optional | recommended***) Create a Desktop or Taskbar shortcut to the script and customize the icon for ease of access.
    * To create a shortcut, simply right-click an open space on your Desktop, select "New -> Shortcut".
    * For the path, you will need to specify the path to Powershell, followed by the path to the relevant `sfdc_script.ps1` script. 
        * Eg. `"C:\Program Files\PowerShell\7\pwsh.exe" -WorkingDirectory ~ "<path to your>\sfdc_script.ps1">`
    * You can also change the icon. I've included an icon in this repo.
    * In order to create a Taskbar shortcut, first launch your newly created shortcut, right-click on the now-running powershell instance IN THE TASKBAR, and select "Pin to taskbar"

## Usage
Simply launch the shortcut or script using Powershell and follow the prompts. Easy-peazy-lemon-squeezy.