# ############################################################################################################################
# Author: Chris Alleaume                                                                                                     #
# Purpose: Create SFDC tasks quickly and easily                                                                              #
# ############################################################################################################################

# Variables for settings

# Enter your SSO email address used in SFDC, or your username
$username = 'c_alleaume@dell.com' 
# Customize this shortened task type list (ensuring that they match what's available in your SFDC instance)
# This makes it easier to only display task types that are relevant to your daily tasks
# Pay attention to the ordering and the colors used to display these tasks further down in this script

$taskTypes = @('Internal Meeting - General',
    'Call Customer',
    'Tech WS - Apps & Data',
    'Tech WS - Multi-Cloud', 
    'Demand Generation: Other', 
    'Accelerator WS - Apps & Data', 
    'Accelerator WS - Multi-Cloud')

$default_bgcolor = (get-host).UI.RawUI.BackgroundColor

Function Get-Tasks {
    Param ($fromDate)
    # List today's tasks
    Write-Host "==> Fetching Tasks..." -ForegroundColor Yellow -NoNewline
    Write-Host " " -BackgroundColor $default_bgcolor -ForegroundColor White
    # Opportunity.Deal_ID__c
    $tasks = sfdx force:data:soql:query -u "$username" --query "SELECT Description, ActivityDate, Account.Name, Type, Workshop_Focus__c, TYPEOF What WHEN Opportunity THEN Deal_ID__c END FROM Task USING SCOPE mine WHERE ActivityDate >= $fromDate ORDER BY ActivityDate DESC" --json | ConvertFrom-Json
    Write-Host ($tasks.result.records | Format-Table -Property @{Name = 'Date'; Expression = { $_.ActivityDate.PadRight(12, ' ') } }, @{Name = 'Deal ID'; Expression = { $_.What.Deal_ID__c + '  ' } }, @{Name = 'Account'; Expression = { $_.Account.Name.subString(0, [System.Math]::Min(20, $_.Account.Name.Length)) + ' ' } },  @{Name = 'Type'; Expression = { $_.Type + '  ' } }, Description | Out-String)
    # $tasks.result.records | Format-List -Property *
}

Write-Host "### ========================================= ###" -ForegroundColor Cyan
Write-Host "### ==========- SFDC Task Creator -========== ###" -ForegroundColor Cyan
Write-Host "### ========================================= ###" -ForegroundColor Cyan
Write-Host " " -BackgroundColor $default_bgcolor

# debug /test code
# $opp = sfdx force:data:get:record -u "$username" -s Task -w "id=00T4v00008o60rQEAQ"
# $task
# end debug / test


$repeat = $true

while ($repeat -eq $true) {
    Write-Host "What do you want to do?" -NoNewLine -ForegroundColor Yellow -BackgroundColor DarkGreen
    Write-Host " " -BackgroundColor $default_bgcolor
    Write-Host "1. Create a Task" -ForegroundColor White
    Write-Host "2. List Tasks created today" -ForegroundColor White
    Write-Host "3. List Tasks created for the past 5 days" -ForegroundColor White
    Write-Host " " -BackgroundColor $default_bgcolor
    $next = $(Write-Host "Choose an action (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
  
    switch ("$next") {
        1 { $repeat = $false; Break }
        2 {
            # List today's tasks
            $act_date = (Get-Date).ToString("yyyy-MM-dd")
            Get-Tasks -fromDate $act_date 
            Break
        }
        3 {
            # List last 5 days' tasks
            $act_date = (Get-Date).AddDays(-5).ToString("yyyy-MM-dd")
            Get-Tasks -fromDate $act_date 
            Break
        }
    }
}


$repeat = $true
$same_opp = $false;
while ($repeat -eq $true) {
    if ($same_opp -eq $false) {
        # Prompt for Deal / Opportunity ID
        $dealId = $(Write-Host "Enter Deal ID: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host).Trim()
        $isDealIdValid = $false
        while ($isDealIdValid -ne $true) {
            if ($dealId -match "^\d+$") {
                $isDealIdValid = $true
            }
            else {
                Write-Host "ERROR: Deal ID [$dealId] is not valid. Please use a valid numeric Deal ID." -ForegroundColor Red
                $dealId = $(Write-Host "Enter Deal ID: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host).Trim()
            }
        }

        # Find opp ID by Deal ID
        Write-Host "==> Fetching Opportunity..." -ForegroundColor Yellow
        $opp = sfdx force:data:soql:query -u "$username" --query "SELECT ID, Name, AccountId, Account.Name FROM Opportunity WHERE Deal_ID__c='$dealId'" --json | ConvertFrom-Json

        # Verify that the Opp was found
        if ($opp.result.totalSize -lt 1) {
            Write-Host "ERROR: Opportunity with Deal ID [$dealId] NOT found (or you don't have access)!" -ForegroundColor Red
            Read-Host -Prompt "Press Enter to exit"
            exit
        }
        Write-Host "==> Opportunity found..." -ForegroundColor Yellow
        # Set variables for future queries
        $oppId = $opp.result.records.Id
        $oppName = $opp.result.records.Name

        # ========================== Removed Account Query ==========================
        # No longer need to query account object - we are now getting this straight from the previous opp query
        # $accountId = $opp.result.records.AccountId
        # Write-Host "==> Fetching Account Details..." -ForegroundColor Yellow
        # Find the associated account so that we can display for verification
        # $account = sfdx force:data:soql:query -u "$username" --query "SELECT Id, Name FROM Account WHERE Id='$accountId'" --json | ConvertFrom-Json
        # $accountName = $account.result.records.Name
        # ================================== End =====================================

        Write-Host " " -BackgroundColor $default_bgcolor
        $accountName = $opp.result.records.Account.Name
    }

    # Retrieve full opp Record by ID (if we need specific details)
    # $opp = sfdx force:data:record:get -s Opportunity -i "$oppId" -u "$username" --json | ConvertFrom-Json
    Write-Host "Account Name: " -ForegroundColor Blue -NoNewline
    Write-Host "$accountName" -ForegroundColor Yellow
    Write-Host "Opportunity Name: " -ForegroundColor Blue -NoNewline
    Write-Host "$oppName" -ForegroundColor Yellow
    Write-Host "Please verify account and opportunity are correct!" -ForegroundColor DarkYellow
    Write-Host " " -BackgroundColor $default_bgcolor

    Write-Host "Choose Task Type:" -NoNewLine -ForegroundColor Yellow -BackgroundColor DarkGreen
    Write-Host " " -BackgroundColor $default_bgcolor
    # Display the task types in two colors to immediately differentiate from common vs less common task types
    Write-Host "1. $($taskTypes[0])" -ForegroundColor White 
    Write-Host "2. $($taskTypes[1])" -ForegroundColor White
    Write-Host "3. $($taskTypes[2])" -ForegroundColor Yellow
    Write-Host "4. $($taskTypes[3])" -ForegroundColor Yellow
    Write-Host "5. $($taskTypes[4])" -ForegroundColor White
    Write-Host "6. $($taskTypes[5])" -ForegroundColor Yellow
    Write-Host "7. $($taskTypes[6])" -ForegroundColor Yellow

    Write-Host " " -BackgroundColor $default_bgcolor
    $taskId = $(Write-Host "Enter Task Type ID (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) 
    $taskType = $taskTypes[$taskId - 1]

    Write-Host "Selected Task Type: $taskType" -ForegroundColor Yellow -BackgroundColor $default_bgcolor

    $wsFocus = ""
    if ($taskType.Contains(" WS ")){
        Write-Host " " -BackgroundColor $default_bgcolor
        Write-Host "Choose a Workshop Focus:" -NoNewLine -ForegroundColor Yellow -BackgroundColor DarkGreen
        Write-Host " " -BackgroundColor $default_bgcolor
        Write-Host "1. Technical WS - Multi-Cloud" -ForegroundColor White 
        Write-Host "2. Technical WS - Apps & Data" -ForegroundColor White 
        Write-Host "3. Chan/GA/OEM Strategy WS" -ForegroundColor White
        Write-Host " " -BackgroundColor $default_bgcolor
        $wsId = $(Write-Host "Enter Workshop Focus ID (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
        switch ("$wsId") {
            1 { 
                $wsFocus = "Technical WS - Multi-Cloud"
                Break 
            }
            2 {
                $wsFocus = "Technical WS - Apps & Data"
                Break 
            }
            3 {
                $wsFocus = "Chan/GA/OEM Strategy WS"
                Break 
            }
        }
    }

    Write-Host " " -BackgroundColor $default_bgcolor
    $descr = $(Write-Host "Enter Task Description: " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)

    $now = Get-Date -Format "yyyy-MM-dd"
    
    $nowTime = Get-Date -Format s
    Write-Host " " -BackgroundColor $default_bgcolor

    # Establish whether or not to use today's date for the task
    $useToday = $(Write-Host "Use today's date for Task? (y/n): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host) 

    if ($useToday.ToLower() -eq 'n') {
        $inDate = Read-Host -Prompt "Enter Task Date in the format [MM/dd/yy] (eg. 10/31/23)"
        $customDate = [datetime]::ParseExact($inDate, 'MM/dd/yy', $null)
        $now = $customDate.ToString("yyyy-MM-dd")
        Write-Host "Task Date: $now" -ForegroundColor Blue
        $nowTime = $customDate.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ss.fffffffZ" )
    } 
    Write-Host " " -BackgroundColor $default_bgcolor

    Write-Host "==> Creating Task [$taskType]..." -ForegroundColor Yellow -BackgroundColor $default_bgcolor
    Write-Host " " -BackgroundColor $default_bgcolor
    # Create the task and display a summary
    $task = sfdx force:data:record:create -s Task -v "Subject='Services' Description='$descr' Status='Completed' Priority='Normal' Workshop_Focus__c='$wsFocus' ActivityDate='$now' WhatId=$oppId IsReminderSet=false TaskSubtype='Task' Type='$taskType' ReminderDateTime='$nowTime'" -u "$username" --json | ConvertFrom-Json
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
    }
    else {
        Write-Host "-################  ERROR  #################-" -ForegroundColor Red
        Write-Host "-- There was an error creating your task! --" -ForegroundColor Red
        Write-Host "------------- Please try again -------------" -ForegroundColor Red
        Write-Host "-##########################################-" -ForegroundColor Red
    }
    Write-Host " " -BackgroundColor $default_bgcolor
    Write-Host "==> What's next?" -NoNewLine -ForegroundColor Yellow -BackgroundColor DarkGreen
    Write-Host " " -BackgroundColor $default_bgcolor
    Write-Host "1. Create a task for a " -NoNewline -ForegroundColor White 
    Write-Host "different " -ForegroundColor Yellow -NoNewline
    Write-Host "opportunity" -ForegroundColor White
    Write-Host "2. Create another task for the same Opportunity" -ForegroundColor White
    Write-Host "3. Open the recently created SFDC task in a browser" -ForegroundColor White
    Write-Host "4. Quit" -ForegroundColor White
    Write-Host " " -BackgroundColor $default_bgcolor
    $next = $(Write-Host "Choose an action (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
    Write-Host " " -BackgroundColor $default_bgcolor
    $repeat = $false
    $same_opp = $false;
    switch ("$next") {
        3 {
            Write-Host "### ============  Opening Task in SFDC  ============ ###" -ForegroundColor Yellow;
            Start-Sleep -Seconds 1
            Start-Process "https://dell.lightning.force.com/lightning/r/Task/$($task.result.Id)/view?ws=%2Flightning%2Fr%2FOpportunity%2F$oppId%2Fview";
            exit;
        }
        2 { $repeat = $true; $same_opp = $true; Break }
        1 { $repeat = $true; $same_opp = $false; Break }
    }
}
Write-Host "### ===============  Done  =============== ###" -ForegroundColor Green
Start-Sleep -Seconds 1.2
#$task = sfdx force:data:record:get -s Task -i "$taskId" -u "$username" --json | ConvertFrom-Json


