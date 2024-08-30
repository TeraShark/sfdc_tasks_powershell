Write-Host "Setting SFDX CLI variables..."
Set-Item -Path Env:SF_AUTOUPDATE_DISABLE -Value $false
Set-Item -Path Env:SFDX_HIDE_RELEASE_NOTES -Value $true
Set-Item -Path Env:SFDX_HIDE_RELEASE_NOTES_FOOTER -Value $true
Write-Host "Settings applied..."


$SiteUrl = "https://dell.sharepoint.com/sites/Pearlj1tech-Team"
$ListName = "SP Tasks"
$NewSiteUrl = "https://dell.sharepoint.com/sites/DevOpsCloud-NativeSPEAR"

$default_bgcolor = (get-host).UI.RawUI.BackgroundColor

$SPConnection = $null
$NewSPConnection = $null

Write-Host "Connecting to Source Sharepoint Site..." -ForegroundColor Yellow

    #Connect to PnP Online if above fails
    $SPConnection = Connect-PnPOnline -Url $SiteURL -UseWebLogin -ReturnConnection

Write-Host "=> Connection successful..." -ForegroundColor Green

Write-Host " " -ForegroundColor White

Write-Host "Connecting to Target Sharepoint Site..." -ForegroundColor Yellow

    #Connect to PnP Online if above fails
    $NewSPConnection = Connect-PnPOnline -Url $NewSiteUrl -UseWebLogin -ReturnConnection
    $web = Get-PnPWeb -Connection $NewSPConnection
 
    
Write-Host "=> Connection successful..." -ForegroundColor Green
Write-Host "====================================================" -ForegroundColor White


Write-Host " "

$Counter = 0
Write-Host "Fetching Source View Columns..." -ForegroundColor White
$view = Get-PnPView -List $ListName -Identity "All Items" -Connection $SPConnection
$columns = $view.ViewFields
#Write-Host "Retrieved $($columns.Length) Columns..."

Write-Host "Fetching Source List Items..." -ForegroundColor White
$listItems = Get-PnPListItem -Connection $SPConnection -List $ListName -PageSize 2000
Write-Host "Parsing Source List and Importing to Target List..."
foreach ($item in $listitems){
    $hashTable=@{}
    #$obj=New-Object -TypeName PSObject
    #$HashTable += $obj
    $title = ""
    Write-Host "Building Target List Item..." -ForegroundColor White
    foreach($column in $columns){
        # Write-Host "Column: $column ||" -ForegroundColor White -NoNewline
        # if ($item.FieldValues[$column] -ne $null){
        #     Write-Host "Value: $($item.FieldValues[$column].toString())" -ForegroundColor White
        # } else {
        #     Write-Host "** Value: NULL **" -ForegroundColor Yellow
        # }
        if($column -eq "LinkTitle"){    
            #$obj | Add-Member -MemberType NoteProperty -Name Title -Value $item.FieldValues.Title
            $hashTable["Title"] = $item.FieldValues.Title
        }else{
            if($item.FieldValues[$column] -ne $null -and $item.FieldValues[$column].toString() -eq "Microsoft.SharePoint.Client.FieldLookupValue"){
                #$obj | Add-Member -MemberType NoteProperty -Name $column -Value $item.FieldValues[$column].LookupValue
                $hashTable["$column"] = $item.FieldValues[$column].LookupValue
            } elseif ($item.FieldValues[$column] -ne $null -and $item.FieldValues[$column].toString() -eq "Microsoft.SharePoint.Client.FieldUrlValue ") {
                $hashTable["$column"] = $item.FieldValues[$column].Url
            }
            elseif ($item.FieldValues[$column] -ne $null -and $item.FieldValues[$column].toString() -eq "Microsoft.SharePoint.Client.FieldUserValue") {
                $user = $web.EnsureUser($item.FieldValues[$column].Email)
                $ctx = Get-PnPContext -Connection $NewSPConnection
                $ctx.Load($user)
                try{
                    $ctx.ExecuteQuery()
                    $hashTable["$column"] = $item.FieldValues[$column].Email
                }
                catch{
                    Write-Host "** User [$($item.FieldValues[$column].Email)] is no longer valid and will not be imported" -ForegroundColor Red
                }
            } elseif ($item.FieldValues[$column] -ne $null -and $item.FieldValues[$column].toString() -eq "Microsoft.SharePoint.Client.FieldUserValue[]") {
                $userValueCollection = [Microsoft.SharePoint.Client.FieldUserValue[]]$item[$column]
                $ids = New-Object System.Collections.ArrayList
                foreach ($FieldUserValue in $userValueCollection)
                {
                    $user = $web.EnsureUser($FieldUserValue.Email)
                    $ctx = Get-PnPContext -Connection $NewSPConnection
                    $ctx.Load($user)
                    try{
                        $ctx.ExecuteQuery()
                        $ids.Add($FieldUserValue.Email)
                    }
                    catch{
                        Write-Host "** User [$($FieldUserValue.Email)] is no longer valid and will not be imported" -ForegroundColor Red
                    }
              
                }
                #Write-Host " "
                $ids_array = $ids.ToArray()
                # $hashTable["$column"] = $ids_array -join ","
                $hashTable["$column"] = $ids_array
            }
            else{
                #$obj | Add-Member -MemberType NoteProperty -Name $column -Value $item.FieldValues[$column]
                $hashTable["$column"] = $item.FieldValues[$column]
            }
        }
    }
    $title = $item.FieldValues.Title
    Write-Host "Importing item [$title]..." -ForegroundColor White
    try{
        Add-PnPListItem -List $ListName -Connection $NewSPConnection -Values $hashTable
    } catch{
        $hashTable
        Write-Host "ERROR: $($Error[0])" -ForegroundColor Red
        $NewSPConnection = $null
        $SPConnection = $null
        Break
    }
   } 

   Write-Host "Done..." -ForegroundColor White
   $NewSPConnection = $null
   $SPConnection = $null

#$camlQuery = "<View><Query><Where><And><Neq><FieldRef Name='Status'/><Value Type='Choice'>Cancelled</Value></Neq><And><Neq><FieldRef Name='Status'/><Value Type='Choice'>Lost</Value></Neq><And><IsNotNull><FieldRef Name='SFDC_x0020_ID'/></IsNotNull><And><Neq><FieldRef Name='Status'/><Value Type='Choice'>Complete</Value></Neq><Neq><FieldRef Name='Status'/><Value Type='Choice'>Cancelled / Archived</Value></Neq></And></And></And></And></Where></Query><ViewFields><FieldRef Name='ID' /><FieldRef Name='Title' /><FieldRef Name='Description' /><FieldRef Name='SFDC_x0020_ID' /><FieldRef Name='TCV_x0020__x0024_' /><FieldRef Name='Est_x002e__x0020_Close' /><FieldRef Name='SFDCLink' /><FieldRef Name='PrimaryContact' /></ViewFields></View>"
