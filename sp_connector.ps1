#region Environment Setup

Set-Item -Path Env:SF_AUTOUPDATE_DISABLE -Value $true
Set-Item -Path Env:SF_HIDE_RELEASE_NOTES -Value $true
Set-Item -Path Env:SF_HIDE_RELEASE_NOTES_FOOTER -Value $true
Set-Item -Path Env:SF_SKIP_NEW_VERSION_CHECK -Value $true
Set-Item -Path Env:SF_DISABLE_TELEMETRY -Value $true
Set-Item -Path Env:PNPPOWERSHELL_UPDATECHECK -Value 'Off'
# [System.Environment]::SetEnvironmentVariable('PNPPOWERSHELL_UPDATECHECK', 'Off')

#endregion

$default_bgcolor = (get-host).UI.RawUI.BackgroundColor

Import-Module "$PSScriptRoot\utils.psm1"

Install-ModuleIfNeeded -Name "PnP.PowerShell" -MaxVersion "2.99"
Install-ModuleIfNeeded -Name "powershell-yaml"
Import-Module -Name "powershell-yaml"

$settingsPath = "$PSScriptRoot\settings.yml"
$yamlSettings = Get-YamlSettings -Path $settingsPath

#region Prechecks

Write-Host "=> Beginning prechecks..." -ForegroundColor Cyan

$lastUpdate = [DateTime]::ParseExact($yamlSettings.settings['last-update'], 'MM/dd/yyyy HH:mm:ss', $null)
# Write-Output $lastUpdate
$elapsed = New-TimeSpan -Start $lastUpdate -End $(Get-Date).ToString()
$updateFreq = $yamlSettings.settings['update-frequency-days']
if($elapsed.Days -ge $updateFreq)
{
    Write-Host "=> Checking and Updating SalesForce CLI..." -ForegroundColor Cyan
    sf update > $null
    # Write-Output $(Get-Date -Format "MM/dd/yyyy HH:mm:ss").ToString()
    $yamlSettings.settings['last-update'] = $(Get-Date -Format "MM/dd/yyyy HH:mm:ss").ToString()
    Save-YamlSettings -YamlSettings $yamlSettings -Path $settingsPath
}
else {
    Write-Host "=> Skipping SalesForce CLI update. Update frequency: $updateFreq days. Remaining days: $($updateFreq - $elapsed.Days)."
}

Write-Host " " 
Write-Host "========== Testing Sharepoint connection ===========" -ForegroundColor White
# Load Shareoiunt settings from YAML
$SiteUrl = $yamlSettings.settings['sharepoint-tracker-url']
$ListName = $yamlSettings.settings['sharepoint-tracker-listname']

$SelectedFields = $yamlSettings.settings['sharepoint-tracker-default-fields']

$SPConnection = $null
# Check for username stored in config file, and if non-existent, prompt and create file
Write-Host "Checking user configuration..."
$username = Get-UserName -ConfigPath "$PSScriptRoot\user.cfg"
$SPConnection = Connect-SharePoint -SiteUrl $SiteUrl

$showSummary = [System.Convert]::ToBoolean($yamlSettings.settings['show-summary'])

Write-Host "=> Prechecks complete. Loading menu..." -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor White
Start-Sleep -Seconds 1
Clear-Host
Write-Host " "

#endregion

Function PostSyncRoutine {
    Write-Host "=> Press Enter to Synchronize Sharepoint SP Tasks to SFDC, or CTRL+C to exit..." -ForegroundColor White
    Read-Host
    & "$PSScriptRoot\sfdc_sync_tasks.ps1"
    Exit
}


Write-Host "Which Sharepoint Tracker Opportunities do you want to sync?" -NoNewLine -ForegroundColor Yellow -BackgroundColor DarkGreen
Write-Host " " -BackgroundColor $default_bgcolor
Write-Host "1. All Active Opportunities" -ForegroundColor White
Write-Host "2. Active Opportunities created Today" -ForegroundColor White
Write-Host "3. Active Opportunities created in the last 5 days" -ForegroundColor White
Write-Host "4. Active Opportunities created by Me" -ForegroundColor White
Write-Host "5. Single Opportunity by SFDC Deal ID" -ForegroundColor White

Write-Host " " -BackgroundColor $default_bgcolor
$next = $(Write-Host "Choose an action (eg 1): " -ForegroundColor Yellow -BackgroundColor DarkGreen -NoNewLine; Read-Host)

Add-Type -Path "$PSScriptRoot\XmlTools\XmlTools.cs"
$xmlTools = [XmlTools.XmlTools]::new($PSScriptRoot)

$camlQuery = $xmlTools.getCAML([XmlTools.XmlTools+CAMLType]::Default)

switch ("$next") {
    1 { Break }
    2 {
        # Today's items
        $camlQuery = $xmlTools.getCAML([XmlTools.XmlTools+CAMLType]::Today)
        $showSummary = $false
        Break
    }
    3 {
        # Last 5 days' items
        $today = (Get-Date)
        $act_date = $today.AddDays(-5).ToString("yyyy-MM-dd")
        $camlQuery = $xmlTools.getCAML([XmlTools.XmlTools+CAMLType]::SinceDate, $act_date)
        Break
    }
    4 {
        # My items
        $camlQuery = $xmlTools.getCAML([XmlTools.XmlTools+CAMLType]::MyItems)
        $showSummary = $true
        Break
    }
    5 {
        # By SFDC Deal ID
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
        Write-Host " " -BackgroundColor $default_bgcolor
        $camlQuery = $xmlTools.getCAML([XmlTools.XmlTools+CAMLType]::ByDealID, $dealId)
        $showSummary = $false
        Break
    }
}


Write-Host " "
Write-Host "Fetching Sharepoint Tracker Opportunities..." -ForegroundColor White
$ListItems = Get-PnPListItem -List $ListName -Query $camlQuery -Connection $SPConnection

if ($null -eq $ListItems) {
    Write-Host "======================================================================" -ForegroundColor Yellow
    Write-Host "  - No Sharepoint Tracker Opportunities found matching your criteria." -ForegroundColor Yellow
    Write-Host "======================================================================" -ForegroundColor Yellow
    PostSyncRoutine
}

Write-Host "Retrieved " -NoNewline
Write-Host "$($ListItems.Count)" -ForegroundColor Green -NoNewline
Write-Host " Sharepoint Tracker Opportunites..." -ForegroundColor White
Write-Host " "

#region Variable declarations

$Counter = 0
$itemsUpdated = 0
$lostItems = @()
$wonItems = @()
$invalidItems = @()
$skippedItems = @()
$sixties = @()
$nineties = @()

#endregion

#region Loop through Sharepoint Tracker Opportunities
$ListItems | ForEach-Object {
    $ListItem = Get-PnPProperty -Connection $SPConnection -ClientObject $_ -Property FieldValuesAsText
    $ListRow = New-Object PSObject
    $Counter++

    ForEach ($Field in $SelectedFields) {
        $ListRow | Add-Member -MemberType NoteProperty $Field $ListItem[$Field]
    }
        
    Write-Progress -PercentComplete ($Counter / $($ListItems.Count) * 100) -Activity "Syncing Opportunities against SFDC..." -Status  "Processing Opportunity $Counter of $($ListItems.Count)"
    # Find opp ID by Deal ID
    $dealId = $($ListRow.SFDC_x0020_ID)
    $cust = $($ListRow.Title)
    $oppDesc = $($ListRow.Description)
    Write-Host "==============================================-->> " -NoNewLine -ForegroundColor White
    Write-Host "Begin [$dealId]" -NoNewline -ForegroundColor Yellow
    Write-Host " <<--==============================================" -ForegroundColor White

    # Verify non-DFN opp (no access to DFN):
    if ($($dealId.Trim().ToLower().StartsWith("dfn"))) {
        Write-Host "--> Note: Skipping DFN Opportunity [$dealId] [$cust - $oppDesc] --> (No DFN access)..." -ForegroundColor Yellow
        Write-Host "==----------------------------------------------##  End [$dealId]  ##----------------------------------------------==" -ForegroundColor Gray
        Write-Host "                                                          |" -ForegroundColor White
        $skippedItems += "[$dealId] [$cust - $oppDesc]"
    }
    else {
        Write-Host "==> Searching for SFDC Opportunity [$dealId] [$cust - $oppDesc]..." -ForegroundColor Cyan
        $start = (Get-Date)
   
        # #=================== Start DEBUG: SFDC JSON ===================

        # Write-Host  "======================  Debug: Start  ======================" # (Debug)
        # Write-Host "Getting Opp JSON from SFDC..." # (Debug)
        $oppJson = sfdx force:data:soql:query -o "$username" --query "SELECT ID, Name, Account.Name, AccountId, Amount, Unweighted_Rev_Services_Only__c, Weighted_Rev_Services_Only__c, Fiscal_Book_Date__c, StageName, Probability, NextStep, Services_Comments__c, Decision_Process__c, Decision_Criteria__c, Metrics__c, Identify_Value_Drivers__c, Campaign__c, Services_Sales_Owner__r.Name, SP_Name__r.Name, CloseDate, IsWon, IsClosed FROM Opportunity WHERE Deal_ID__c='$($dealId.Trim())'" --json
        # $oppJson # (Debug)
        # Write-Host "Parsing JSON from SFDC into PowerShell object..." # (Debug)
        $opp = $oppJson | ConvertFrom-Json
    
        # Write-Host "Parsed JSON from SFDC successfully..." # (Debug)
        # Write-Host  "======================  Debug: End  ======================" # (Debug)

        # #==================== End DEBUG: SFDC JSON ====================

        $finish = (Get-Date)
        Write-Host ">>> Query Time: $(New-TimeSpan -Start $start -End $finish)" -ForegroundColor DarkGray
  
        # Verify that the Opp was found
        if ($opp.result.totalSize -lt 1) {
            Write-Host "ERROR: Opportunity [$dealId] [$cust - $oppDesc] NOT found (or you don't have access)!" -ForegroundColor Red
            # Commented out updating the Notes field (below) on Sharepoint, since this script may be run by multiple people without access to certain records.
            #Set-PnPListItem -List $ListName -Identity $ListRow.ID -Values @{"SFDCNotes" = "** INACCESSIBLE SFDC ID **" } > $null
            $invalidItems += "[$dealId] [$cust - $oppDesc]"
            Write-Host "==----------------------------------------------##  End [$dealId]  ##----------------------------------------------==" -ForegroundColor Gray
            Write-Host "                                                          |" -ForegroundColor White
        }
        else {
            $oppName = $opp.result.records.Name
            $oppId = $opp.result.records.ID
            $amount = '{0:C}' -f $opp.result.records.Amount
            $servicesAmount = '{0:C}' -f $opp.result.records.Unweighted_Rev_Services_Only__c
            if (!$servicesAmount) {
                $servicesAmount = '{0:C}' -f $opp.result.records.Weighted_Rev_Services_Only__c
            }
            $estCloseDate = $opp.result.records.Fiscal_Book_Date__c
            $dateClosed = $opp.result.records.CloseDate
            $isWon = ($opp.result.records.IsWon -eq 'true' -and $opp.result.records.IsClosed -eq 'true')

            $stage = $opp.result.records.StageName
            $prob = $opp.result.records.Probability
            $notes = ""
            $accName = $opp.result.records.Account.Name

            Write-Host "==> SFDC Opportunity [$dealId] ($oppName) found..." -ForegroundColor Cyan
            Write-Host "==> Account: $accName" -ForegroundColor Cyan

            $($opp.result.records.Total_Contract_Value_TCV__c)
            if ($servicesAmount) {
                Write-Host "==> TCV: $amount | Services: $servicesAmount | Stage: $stage" -ForegroundColor White
            }
            else {
                Write-Host "==> TCV: $amount | Services: <blank> | Stage: $stage" -ForegroundColor Cyan
            }
            
            if (![string]::IsNullOrEmpty($opp.result.records.Services_Comments__c)) {
                $notes = "Comments: " + $opp.result.records.Services_Comments__c + "`r`n"
            }

            if (![string]::IsNullOrEmpty($opp.result.records.NextStep)) {
                $notes += "Next Step: " + $opp.result.records.NextStep + "`r`n"
            }

            if (![string]::IsNullOrEmpty($opp.result.records.Decision_Criteria__c)) {
                $notes += "Decision Criteria: " + $opp.result.records.Decision_Criteria__c + "`r`n"
            }
            
            if (![string]::IsNullOrEmpty($opp.result.records.Services_Sales_Owner__r.Name)) {
                $notes += "SAE: " + $opp.result.records.Services_Sales_Owner__r.Name + "`r`n"
            }
            
            if (![string]::IsNullOrEmpty($opp.result.records.SP_Name__r.Name)) {
                $notes += "Primary SP: " + $opp.result.records.SP_Name__r.Name + "`r`n"
            }

            $updates = @{"TCV_x0020__x0024_" = "$amount"; "Est_x002e__x0020_Close" = "$estCloseDate"; "SFDC_Internal_ID" = "$oppId"; "SFDCLink" = "https://dell.lightning.force.com/lightning/r/Opportunity/$oppId/view"; "SFDCStage" = "$stage"; "SFDCProb" = "$prob"; "SFDCNotes" = "$notes" }
            $updates += @{"SFDC_x0020_Account_x0020__x0026_" = "Account: $accName`r`nOpp: $oppName" }
            if ($servicesAmount) {
                $updates += @{"Value" = "$servicesAmount" } 
            }
            if ($dateClosed -and $isWon) {
                $updates += @{"Close_x0020_Date" = "$dateClosed" }
            }
            #MEDDPIC validation:
            if ([string]::IsNullOrEmpty($opp.result.records.Decision_Process__c) -and 
                [string]::IsNullOrEmpty($opp.result.records.Decision_Criteria__c) -and 
                [string]::IsNullOrEmpty($opp.result.records.Metrics__c) -and 
                [string]::IsNullOrEmpty($opp.result.records.Identify_Value_Drivers__c)) {
                $updates += @{"MEDDPIC" = "Not Compliant" }
            }
            elseif (![string]::IsNullOrEmpty($opp.result.records.Decision_Process__c) -and 
                ![string]::IsNullOrEmpty($opp.result.records.Decision_Criteria__c) -and 
                ![string]::IsNullOrEmpty($opp.result.records.Metrics__c) -and 
                ![string]::IsNullOrEmpty($opp.result.records.Identify_Value_Drivers__c)) {
                $updates += @{"MEDDPIC" = "Compliant" }
            }
            #Opportunity Lost:
            if ($stage.StartsWith("Lost")) {
                $updates += @{"Status" = "Lost"; "Booked" = "0" }
                Write-Host "** Note: Opportunity [$dealId] was Lost..." -ForegroundColor Red
                $lostItems += "[$dealId] [$cust - $oppDesc] TCV: $amount"
            }
            elseif ($stage.StartsWith("Win")) {
                $updates += @{"Booked" = "1" }
                Write-Host "** Note: Opportunity [$dealId] has been WON..." -ForegroundColor Green
                $wonItems += "[$dealId] [$cust - $oppDesc] TCV: $amount"
            }
            else {
                $updates += @{"Booked" = "0" }
            }
            # 60/90 opps:
            if ($prob -eq 60) {
                $sixties += "$cust ($oppDesc)`n     [$dealId] | TCV: $amount | $estCloseDate"
            }
            elseif ($prob -eq 90) {
                $nineties += "$cust ($oppDesc)`n     [$dealId] | TCV: $amount | $estCloseDate"
            }

            Set-PnPListItem -List $ListName -Identity $ListRow.ID -Values $updates -Connection $SPConnection > $null
            Write-Host "==> Sharepoint Opportunity [$dealId] updated..." -ForegroundColor Green
            $itemsUpdated++
        
            Write-Host "==----------------------------------------------##  End [$dealId]  ##----------------------------------------------==" -ForegroundColor Gray
            Write-Host "                                                          |" -ForegroundColor White
        }
    }
}
#endregion

Write-Host "                                                         ~~~" -ForegroundColor White
#region Summary
if ($showSummary -eq $true) {
    Write-Host "=========================================  SUMMARY  ====================================================" -ForegroundColor Gray
    Write-Host "==> Total Opportunities: $($ListItems.Count) || Opportunities synchronized: " -NoNewline -ForegroundColor White 
    Write-Host "$itemsUpdated " -NoNewline -ForegroundColor Green
    Write-Host "|| Opportunities not synchronized: " -ForegroundColor White -NoNewline
    Write-Host "$($($ListItems.Count) - $itemsUpdated)" -ForegroundColor Red
    Write-Host " "
    Write-Host "----------------------------------- Skipped Items [DFN] ($($skippedItems.count)) --------------------------------------------" -ForegroundColor Gray
    foreach ( $item in $skippedItems) {
        Write-Host "  - $item" -ForegroundColor Gray
    }
    Write-Host "--------------------------------------------------------------------------------------------------------" -ForegroundColor Gray
    Write-Host "------------------------------------- Opps Not Found ($($invalidItems.count)) -----------------------------------------------" -ForegroundColor Yellow
    foreach ( $item in $invalidItems) {
        Write-Host "  - $item" -ForegroundColor Yellow
    }
    Write-Host "--------------------------------------------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "----------------------------------- Lost Opportunities ($($lostItems.count)) ---------------------------------------------" -ForegroundColor Red
    foreach ( $item in $lostItems) {
        Write-Host "  - $item" -ForegroundColor Red
    }
    Write-Host "--------------------------------------------------------------------------------------------------------" -ForegroundColor Red
    Write-Host " - - - - - - - - - - - - - - - - - - - - 60 / 90 - - - - - - - - - - - - - - - - - - - - - - - - - - - -" -ForegroundColor Gray
    Write-Host "---------------------------------- Opps @ Propose - 60% ($($sixties.count)) --------------------------------------------" -ForegroundColor Green
    foreach ( $item in $sixties) {
        Write-Host "  ->" -ForegroundColor Green -NoNewline
        Write-Host " $item" -ForegroundColor Gray
    }
    Write-Host " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" -ForegroundColor Green
    Write-Host "---------------------------------- Opps @ Commit - 90% ($($nineties.count)) ---------------------------------------------" -ForegroundColor Green
    foreach ( $item in $nineties) {
        Write-Host "  ->" -ForegroundColor Green -NoNewline
        Write-Host " $item" -ForegroundColor Gray
    }
    Write-Host " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" -ForegroundColor Green
    Write-Host " - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -" -ForegroundColor Gray
    Write-Host "----------------------------------- Won Opportunities ($($wonItems.count)) ----------------------------------------------" -ForegroundColor Green
    foreach ( $item in $wonItems) {
        Write-Host "  - $item" -ForegroundColor Green
    }
    Write-Host "--------------------------------------------------------------------------------------------------------" -ForegroundColor Green
    Write-Host " " -BackgroundColor $default_bgcolor
}
#endregion
Write-Host "==> Process completed on $(Get-Date)" -ForegroundColor White
Write-Host "================================================- DONE -================================================" -ForegroundColor Green
Write-Host " " -BackgroundColor $default_bgcolor

PostSyncRoutine
# exit
