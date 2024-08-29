

Write-Host "Setting SFDX CLI variables..."
Set-Item -Path Env:SF_AUTOUPDATE_DISABLE -Value $false
Set-Item -Path Env:SFDX_HIDE_RELEASE_NOTES -Value $true
Set-Item -Path Env:SFDX_HIDE_RELEASE_NOTES_FOOTER -Value $true
Write-Host "Settings applied..."
# ====================================================== Set Variables =================================================================
$SiteUrl = "https://dell.sharepoint.com/sites/Pearlj1tech-Team"
$ListName = "SP Tasks"
#InternalName of the selected fields
$SelectedFields = @("ID", "Title", "Description", "SFDC_x0020_ID", "Task_x0020_Type", "WS_x0020_Focus", "Task_x0020_Date", "Sync_x0020_Status", "Author", "SFDC_Internal_ID")

$default_bgcolor = (get-host).UI.RawUI.BackgroundColor


Function Save-UserName {
    $username = $(Write-Host "Please enter your email address as it appears in your SFDC Profile:" -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
    Set-Content "$PSScriptRoot\user.cfg" -Value $username
    return $username
}

Function Save-SharepointName {
    $SP_UserName = $(Write-Host 'Please enter your Display Name as it appears in your Sharepoint profile. This is usually in the format "LastName, FirstName":' -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)
    Set-Content "$PSScriptRoot\sharepoint.cfg" -Value $SP_UserName
    return $SP_UserName
}

# Check for username stored in config file, and if non-existent, prompt and create file
$username = ''
if (Test-Path "$PSScriptRoot\user.cfg") {
    $username = Get-Content "$PSScriptRoot\user.cfg"
} else {
    $username = Save-UserName
}
# Validate length of User Name, and re-prompt if invalid
if ($username.Length -lt 8){
    $username = Save-UserName
}

$SP_UserName = '' 
if (Test-Path "$PSScriptRoot\sharepoint.cfg") {
    $SP_UserName = Get-Content "$PSScriptRoot\sharepoint.cfg"
} else {
    $SP_UserName = Save-SharepointName
}
# Validate length of User Name, and re-prompt if invalid
if ($SP_UserName.Length -lt 5){
    $SP_UserName = Save-SharepointName
}

#Check whether the PnP.Powershell module is installed, and install it if not
if (!(Get-Module -ListAvailable -Name "PnP.Powershell")){
    Write-Host "Installing Sharepoint Powershell module..."
    Install-Module -Name "PnP.PowerShell"
    Write-Host "Sharepoint Powershell module installed..."
}

$camlQuery = "<View><Query><Where><Neq><FieldRef Name='Sync_x0020_Status'/><Value Type='Choice'>Synced</Value></Neq></Where></Query><ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Description'/><FieldRef Name='SFDC_x0020_ID'/><FieldRef Name='Task_x0020_Type'/><FieldRef Name='WS_x0020_Focus'/><FieldRef Name='Task_x0020_Date'/><FieldRef Name='Sync_x0020_Status'/><FieldRef Name='Author'/><FieldRef Name='SFDC_Internal_ID'/></ViewFields></View>"
# ====================================================== End Variables ===================================================================
$SPConnection = $null
Write-Host "========== Testing Sharepoint connection ===========" -ForegroundColor White
try {
    # First try to get list item from cached connection
    $SPConnection = Get-PnPConnection | Out-Null
    Write-Host "=> Conection successful..." -ForegroundColor Green
    Write-Host "====================================================" -ForegroundColor White
}
catch {
    Write-Host "Reconnecting to Sharepoint..." -ForegroundColor Yellow
    #Connect to PnP Online if above fails
    $SPConnection = Connect-PnPOnline -Url $SiteURL -UseWebLogin | Out-Null
    Write-Host "=> Conection successful..." -ForegroundColor Green
    Write-Host "====================================================" -ForegroundColor White
}

$Counter = 0

#PageSize:The number of items to retrieve per page request
#$ListItems = Get-PnPListItem -List $ListName -Fields $SelectedFields 
$ListItems = Get-PnPListItem -List $ListName -Query $camlQuery -Connection $SPConnection
Write-Host "Retrieved $($ListItems.Count) UNSYNCHRONIZED SP Task(s)..."
#Get all items from list
$itemsUpdated = 0

$ListItems | ForEach-Object {
    $ListItem = Get-PnPProperty -ClientObject $_ -Property FieldValuesAsText
    $ListRow = New-Object PSObject
    $Counter++
    ForEach ($Field in $SelectedFields) {
        $ListRow | Add-Member -MemberType NoteProperty $Field $ListItem[$Field]
    }
    Write-Progress -PercentComplete ($Counter / $($ListItems.Count) * 100) -Activity "Syncing Tasks to SFDC..." -Status  "Processing Item $Counter of $($ListItems.Count)"
    # Find opp ID by Deal ID
    $dealId = $($ListRow.SFDC_x0020_ID)
    $dealId = $dealId.Trim()
    $cust = $($ListRow.Title)
    $taskType = $($ListRow.Task_x0020_Type)
    $taskDesc = [string]$($ListRow.Description) -replace "`r|`n|`t", ". " -replace "'", ""
    $taskDesc = $taskDesc.Trim()
    $inDate = $($ListRow.Task_x0020_Date)
    $customDate = [datetime]::ParseExact($inDate, 'M/d/yyyy', $null)
    $taskDate = $customDate.ToString("yyyy-MM-dd")
    $author = $($ListRow.Author)
    # $taskDateTime = $customDate.ToUniversalTime().ToString( "yyyy-MM-ddTHH:mm:ss.fffffffZ" )
    $wsFocus = $($ListRow.WS_x0020_Focus)
    

    # ========= Create SFDC Task =========
    Write-Host "============================-->> Begin [$dealId] <<--============================" -ForegroundColor White
    $start = (Get-Date)  
    Write-Host "==> Fetching Opportunity [$dealId] ($cust)..." -ForegroundColor Cyan
    $opp = sfdx force:data:soql:query -u "$username" --query "SELECT ID, Name FROM Opportunity WHERE Deal_ID__c='$dealId'" --json | ConvertFrom-Json
    $finish = (Get-Date)   
    Write-Host ">>> Query Time: $(New-TimeSpan -Start $start -End $finish)" -ForegroundColor Gray
    $oppId = $opp.result.records.Id
    $oppName = $opp.result.records.Name
    
    Write-Host "==> Creating Task [$taskType] for [$author] Opportunity [$dealId] [($cust) - $oppName]..." -ForegroundColor Cyan
    # $username = "c_alleaume@dell.com"
    $ownerId = "0054v00000CnhM7AAJ" # Chris Alleaume's SFDC Internal User ID
    if ($author.Contains("Demarrais")){
        # $username = "karl.demarrais@dell.com"
        $ownerId = "0054v00000EVZRqAAP" # Karl's SFDC Internal User ID (0054v00000EVZRqAAP)
    }

    $start = (Get-Date)
    $safeTaskDesc = [System.Web.HttpUtility]::UrlEncode("$taskDesc")
    if ($wsFocus.Contains("None")){
        Write-Host "Command params: Subject='Services' Description='$taskDesc' Status='Completed' Priority='Normal' ActivityDate='$taskDate' WhatId='$oppId' IsReminderSet=false TaskSubtype='Task' Type='$taskType' OwnerId=$ownerId"
        $result = sfdx force:data:record:create -s Task -v "Subject='Services' Description='$safeTaskDesc' Status='Completed' Priority='Normal' ActivityDate='$taskDate' WhatId='$oppId' IsReminderSet=false TaskSubtype='Task' Type='$taskType' OwnerId=$ownerId ReminderDateTime='$taskDateTime'" -u "$username" --json --loglevel debug
        # $task = sf data create record --sobject Task --values "Subject='Services' Description='$taskDesc' Status='Completed' Priority='Normal' Workshop_Focus__c='$wsFocus' ActivityDate='$taskDate' WhatId=$oppId IsReminderSet=false TaskSubtype='Task' Type='$taskType' OwnerId=$ownerId ReminderDateTime='$taskDateTime'" -u "$username" --json | ConvertFrom-Json    
    } else {
        Write-Host "Command params: Subject='Services' Description='$taskDesc' Status='Completed' Priority='Normal' ActivityDate='$taskDate' WhatId='$oppId' IsReminderSet=false TaskSubtype='Task' Type='$taskType' OwnerId=$ownerId"
        $result = sfdx force:data:record:create -s Task -v "Subject='Services' Description='$safeTaskDesc' Status='Completed' Priority='Normal' Workshop_Focus__c='$wsFocus' ActivityDate='$taskDate' WhatId='$oppId' IsReminderSet=false TaskSubtype='Task' Type='$taskType' OwnerId=$ownerId ReminderDateTime='$taskDateTime'" -u "$username" --json --loglevel debug
        # $result
    }
    
    $task = $result | ConvertFrom-Json
    # Write-Host ($task | Format-Table | Out-String)

    $finish = (Get-Date)   
    Write-Host ">>> Query Time: $(New-TimeSpan -Start $start -End $finish)" -ForegroundColor Gray

    if ($task.result.success -eq 'True') {
        Write-Host "Task [" -ForegroundColor Green -NoNewline
        Write-Host "$taskType" -ForegroundColor Yellow -NoNewline
        Write-Host "] created successfully" -ForegroundColor Green
        Write-Host "Task ID: " -ForegroundColor Green -NoNewline
        Write-Host "$($task.result.Id)" -ForegroundColor Yellow
        Write-Host "Account Name: " -ForegroundColor Green -NoNewline
        Write-Host "$cust" -ForegroundColor Yellow
        Write-Host "Task Description: " -ForegroundColor Green -NoNewline
        Write-Host "$taskDesc" -ForegroundColor Yellow
        Write-Host "Deal ID: " -ForegroundColor Green -NoNewline
        Write-Host "$dealId" -ForegroundColor Yellow

        $taskId = $($task.result.Id)
        $updates = @{"Sync_x0020_Status" = "Synced"; "SFDC_x0020_Task_x0020_Link" = "https://dell.lightning.force.com/lightning/r/Task/$taskId/view" }
       
        Set-PnPListItem -List $ListName -Identity $ListRow.ID -Values $updates -Connection $SPConnection > $null
        Write-Host "==> Sharepoint Task Item for Opp [$dealId] updated..." -ForegroundColor Green
        $itemsUpdated++
    }
    else {
        Write-Host "-###############################  ERROR  #####################################-" -ForegroundColor Red
        Write-Host "-- There was an error creating your task! --" -ForegroundColor Red
        Write-Host "This is usually due to one of the following reasons:" -ForegroundColor Yellow
        Write-Host "1. The Task entry in Sharepoint is missing SFDC connection data." -ForegroundColor Yellow
        Write-Host "   This happens if you created a task against an unsynced Tracker opportunity." -ForegroundColor Yellow
        Write-Host "   To fix this, make sure that the Tracker opportunity is synced against SFDC." -ForegroundColor Yellow
        Write-Host "   Then, delete the Task (from Sharepoint) and recreate it through the Tracker." -ForegroundColor Yellow
        Write-Host "2. You don't have access to the associated SFDC opportunity for this task." -ForegroundColor Yellow
        Write-Host "3. The SFDC Opportunity is closed or no longer exists." -ForegroundColor Yellow
        Write-Host "4. Occasionally, SFDC has a moment, and the API call fails. In this case," -ForegroundColor Yellow
        Write-Host "------------------------------ Please try again ------------------------------" -ForegroundColor Red
        Write-Host "-############################################################################-" -ForegroundColor Red
    }
   
    Write-Host "==----------------------------##  End [$dealId]  ##----------------------------==" -ForegroundColor White

}
Write-Host "===================== Begin SUMMARY ======================" -ForegroundColor Gray
Write-Host "==> Items processed: $($ListItems.Count), Items updated: $itemsUpdated" -ForegroundColor Green
Write-Host "========================== DONE ==========================" -ForegroundColor Gray
Write-Host "*** Note: if you see warnings about SF being out-of-date, open an Administrator Powershell window, and type 'sfdx update'" -ForegroundColor White
Read-Host -Prompt "Press Enter to exit"
exit
