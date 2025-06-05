function Install-ModuleIfNeeded {
    param([string]$Name, [string]$MaxVersion = $null)
    if (!(Get-Module -ListAvailable -Name $Name)) {
        Write-Host "=> Installing $Name module..."
        if ($MaxVersion) {
            Install-Module -Name $Name -MaximumVersion $MaxVersion -AcceptLicense -Force
        } else {
            Install-Module -Name $Name -AcceptLicense -Force
        }
        Write-Host "=> $Name module installed."
    }
}

function Get-YamlSettings {
    param([string]$Path)
    $settings = Get-Content -Path $Path -Raw
    return ConvertFrom-Yaml -Yaml $settings -Ordered
}

function Save-YamlSettings {
    param([object]$YamlSettings, [string]$Path)
    $YamlSettings | ConvertTo-Yaml -Options WithIndentedSequences | Out-File -FilePath $Path -Encoding utf8
}

function Get-UserName {
    param([string]$ConfigPath)
    if (Test-Path $ConfigPath) {
        return Get-Content $ConfigPath
    } else {
        $username = Read-Host "Please enter your email address as it appears in your SFDC Profile"
        Set-Content $ConfigPath -Value $username
        return $username
    }
}

function Connect-SharePoint {
    param([string]$SiteUrl)
    $conn = $null
    try {
        $conn = Get-PnPConnection
        if ($null -eq $conn) {
            $conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection
        }
    } catch {
        $conn = Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ReturnConnection
    }
    return $conn
}