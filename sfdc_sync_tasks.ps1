# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! IMPORTANT: Change your credentails here !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
$username = 'c_alleaume@dell.com' # Your Corporate Email Address
$SP_UserName = "Alleaume, Chris" # The way your name is presented in Sharepoint (usually, this is in the format "LastName, FirstName")
# =====================================================================================================================================
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! End Credentials !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# =====================================================================================================================================

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

# Write-Host "Loading Sharepoint Powershell module..."
# Install-Module -Name "PnP.PowerShell"
# Write-Host "Sharepoint Powershell module loaded..."

$camlQuery = "<View><Query><Where><Neq><FieldRef Name='Sync_x0020_Status'/><Value Type='Choice'>Synced</Value></Neq></Where></Query><ViewFields><FieldRef Name='ID'/><FieldRef Name='Title'/><FieldRef Name='Description'/><FieldRef Name='SFDC_x0020_ID'/><FieldRef Name='Task_x0020_Type'/><FieldRef Name='WS_x0020_Focus'/><FieldRef Name='Task_x0020_Date'/><FieldRef Name='Sync_x0020_Status'/><FieldRef Name='Author'/><FieldRef Name='SFDC_Internal_ID'/></ViewFields></View>"
# ====================================================== End Variables ===================================================================

Write-Host "========== Testing Sharepoint connection ===========" -ForegroundColor White
try {
    # First try to get list item from cached connection
    Get-PnPListItem -List $ListName -Id 6 > $null
    Write-Host "=> Connection successful..." -ForegroundColor Green
    Write-Host "====================================================" -ForegroundColor White
}
catch {
    Write-Host "Reconnecting to Sharepoint..." -ForegroundColor Yellow
    #Connect to PnP Online if above fails
    Connect-PnPOnline -Url $SiteURL -UseWebLogin
    Write-Host "====================================================" -ForegroundColor White
}

$Counter = 0

#PageSize:The number of items to retrieve per page request
#$ListItems = Get-PnPListItem -List $ListName -Fields $SelectedFields 
$ListItems = Get-PnPListItem -List $ListName -Query $camlQuery
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
       
        Set-PnPListItem -List $ListName -Identity $ListRow.ID -Values $updates > $null
        Write-Host "==> Sharepoint Task Item for Opp [$dealId] updated..." -ForegroundColor Green
        $itemsUpdated++
    }
    else {
        Write-Host "-################  ERROR  #################-" -ForegroundColor Red
        Write-Host "-- There was an error creating your task! --" -ForegroundColor Red
        Write-Host "------------- Please try again -------------" -ForegroundColor Red
        Write-Host "-##########################################-" -ForegroundColor Red
    }
   
    Write-Host "==----------------------------##  End [$dealId]  ##----------------------------==" -ForegroundColor White

}
Write-Host "===================== Begin SUMMARY ======================" -ForegroundColor Gray
Write-Host "==> Items processed: $($ListItems.Count), Items updated: $itemsUpdated" -ForegroundColor Green
Write-Host "========================== DONE ==========================" -ForegroundColor Gray
Write-Host "*** Note: if you see warnings about SFDX being out-of-date, open an Administrator Powershell window, and type 'sfdx update'" -ForegroundColor White
Read-Host -Prompt "Press Enter to exit"
exit
