# ############################################################################################################################
# Author: Chris Alleaume                                                                                                     #
# Purpose: Create SFDC tasks quickly and easily                                                                              #
# ############################################################################################################################

# Variables for settings

# Enter your SSO email address used in SFDC, or your username
$username = 'your_corporate_email_address' 
# Customize this shortened task type list (ensuring that they match what's available in your SFDC instance)
# This makes it easier to only display task types that are relevant to your daily tasks
# Pay attention to the ordering and the colors used to display these tasks further below in this script
$taskTypes = @('Internal Meeting - General',
    'Call Customer',
    'Tech WS - Apps & Data',
    'Tech WS - Multi-Cloud', 
    'Demand Generation: Other', 
    'Accelerator WS - Apps & Data', 
    'Accelerator WS - Multi-Cloud')

Write-Host "### ========================================= ###" -ForegroundColor Cyan
Write-Host "### ==========- SFDC Task Creator -========== ###" -ForegroundColor Cyan
Write-Host "### ========================================= ###" -ForegroundColor Cyan
Write-Host " "

# Prompt for Deal / Opportunity ID
$dealId = $(Write-Host "Enter Deal ID: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host).Trim()
$isDealIdValid = $false
while($isDealIdValid -ne $true){
    if ($dealId -match "^\d+$") {
        $isDealIdValid = $true
    } else {
        Write-Host "ERROR: Deal ID [$dealId] is not valid. Please use a valid numeric Deal ID." -ForegroundColor Red
        $dealId = $(Write-Host "Enter Deal ID: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host).Trim()
    }
}

# Find opp ID by Deal ID
Write-Host "Fetching Opportunity..." -ForegroundColor Yellow
$opp = sfdx force:data:soql:query -u "$username" --query "SELECT ID, Name, AccountId FROM Opportunity WHERE Deal_ID__c='$dealId'" --json | ConvertFrom-Json

# Verify that the Opp was found
if ($opp.result.totalSize -lt 1){
    Write-Host "ERROR: Opportunity with Deal ID [$dealId] NOT found (or you don't have access)!" -ForegroundColor Red
    Read-Host -Prompt "Press Enter to exit"
    exit
}

# Set variables for future queries
$oppId = $opp.result.records.Id
$oppName = $opp.result.records.Name
$accountId = $opp.result.records.AccountId

# Find the associated account so that we can display for verification
$account = sfdx force:data:soql:query -u "$username" --query "SELECT Id, Name FROM Account WHERE Id='$accountId'" --json | ConvertFrom-Json
Write-Host " "
$accountName = $account.result.records.Name

# Retrieve full opp Record by ID (if we need specific details)
# $opp = sfdx force:data:record:get -s Opportunity -i "$oppId" -u "$username" --json | ConvertFrom-Json
Write-Host "Account Name: " -ForegroundColor Blue -NoNewline
Write-Host "$accountName" -ForegroundColor Yellow
Write-Host "Opportunity Name: " -ForegroundColor Blue -NoNewline
Write-Host "$oppName" -ForegroundColor Yellow
Write-Host "Please verify account and opportunity are correct!" -ForegroundColor DarkYellow
Write-Host " "
$now = Get-Date -Format "yyyy-MM-dd"
$nowTime = Get-Date -Format s

# Establish whether or not to use today's date for the task
$useToday = $(Write-Host "Use today's date for Task? (y/n): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) 

if ($useToday -eq 'n'){
    $inDate = Read-Host -Prompt "Enter Task Date in the format [MM/dd/yy] (eg. 10/31/22)"
    $customDate = [datetime]::ParseExact($inDate, 'MM/dd/yy', $null)
    $now = $customDate.ToString("yyyy-MM-dd")
    Write-Host "Task Date: $now" -ForegroundColor Blue
    $nowTime = $customDate.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ss.fffffffZ" )
}
Write-Host " "

Write-Host "Choose Task Type:" -ForegroundColor Yellow -BackgroundColor DarkGreen

# Display the task types in two colors to immediately differentiate from common vs less common task types
Write-Host "1. $($taskTypes[0])" -ForegroundColor White
Write-Host "2. $($taskTypes[1])" -ForegroundColor White
Write-Host "3. $($taskTypes[2])" -ForegroundColor Yellow
Write-Host "4. $($taskTypes[3])" -ForegroundColor Yellow
Write-Host "5. $($taskTypes[4])" -ForegroundColor White
Write-Host "6. $($taskTypes[5])" -ForegroundColor Yellow
Write-Host "7. $($taskTypes[6])" -ForegroundColor Yellow

Write-Host " "
$taskId = $(Write-Host "Enter Task Type ID (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) 
$taskType = $taskTypes[$taskId - 1]

Write-Host "Selected Task Type: $taskType" -ForegroundColor Yellow
Write-Host " "
$descr = $(Write-Host "Enter Task Description: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)

Write-Host "Creating Task [$taskType]..." -ForegroundColor Yellow
Write-Host " "

# Create the task and display a summary
$task = sfdx force:data:record:create -s Task -v "Subject='Services' Description='$descr' Status='Completed' Priority='Normal' ActivityDate='$now' WhatId=$oppId IsReminderSet=false TaskSubtype='Task' Type='$taskType' ReminderDateTime='$nowTime'" -u "$username" --json | ConvertFrom-Json
if ($task.result.success -eq 'True') {
    Write-Host "Task [" -ForegroundColor Green -NoNewline
    Write-Host "$taskType" -ForegroundColor Yellow -NoNewline
    Write-Host "] created successfully" -ForegroundColor Green
    Write-Host "Task ID: " -ForegroundColor Green -NoNewline
    Write-Host "$($task.result.Id)" -ForegroundColor Yellow
    Write-Host "Account Name: " -ForegroundColor Green -NoNewline
    Write-Host "$accountName" -ForegroundColor Yellow
    Write-Host "Opportunity Name: " -ForegroundColor Green -NoNewline
    Write-Host "$oppName" -ForegroundColor Yellow
    Write-Host "Deal ID: " -ForegroundColor Green -NoNewline
    Write-Host "$dealId" -ForegroundColor Yellow
} else {
    Write-Host "There was an error creating your task!" -ForegroundColor Red
}

# Prompt whether to open the newly created task in a browser
$openUrl = $(Write-Host "Open Task in Browser? (y/n): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
if ($openUrl -eq 'y'){
    Write-Host "### ===============  Done  =============== ###" -ForegroundColor Yellow
    Start-Process "https://dell.lightning.force.com/lightning/r/Task/$($task.result.Id)/view?ws=%2Flightning%2Fr%2FOpportunity%2F$oppId%2Fview"
    exit
}
Write-Host "### ===============  Done  =============== ###" -ForegroundColor Yellow
Start-Sleep -Seconds 2
#$task = sfdx force:data:record:get -s Task -i "$taskId" -u "$username" --json | ConvertFrom-Json


