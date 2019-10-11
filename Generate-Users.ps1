param( 
    [Parameter(Mandatory=$true)]
    [String]$Domain,
    
    [Int32]$usercount = 100,

    [String[]]$CompanyName,
    [String[]]$Department,
    [String[]]$JobTitle,
    [String[]]$GivenName,
    [String[]]$Surname
)

$allvaluestext = "City","CompanyName","Department","JobTitle","GivenName","Surname"
$allvalueshash = @{}
$fillinghash = @{}


function Prepare-Execution{
    if(!(Get-Command Connect-azuread)) {
        if(!(([Security.Principal.WindowsPrincipal] ` [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)))
        {
            Write-Host "Please restart the Application as Administrator, you are missing some Modules!" -ForegroundColor RED
            exit(1)
        }
    }
    if(!(Get-Command Connect-azuread)) {
        try{
            Install-Module AzureAD
        }catch{
            Write-Host "Could not install Module : AzureAD. Error: $($error.Exception.message)" -ForegroundColor Red
            exit(1)
        }
    }
}

function Connect-Services {
    $Credentials = Get-Credential -Message "Please enter your Credentials"
    try{
        Connect-AzureAD -Credential $Credentials
    }catch{
        Write-Host "Could not connect to AzureAD! Error: $($error.exception.message)" -Foregroundcolor Red 
    }
}

function Prepare-Excel {
    try{
        $Excel = New-Object -ComObject Excel.Application
        $Scriptdir = Split-Path $script:MyInvocation.MyCommand.Path
        $Workbook = $Excel.Workbooks.Open($Scriptdir + "\data.xlsx")

        $allvaluestext | Foreach-object {
            $valuepos = $workbook.Sheets.Item(1).Cells.Find($_).Address(0,0,1,1).Split('!')[1]
            for($i = 2;$i -lt 999;$i++) {
                $x = [int][char]$valuepos[0] - 64
                $value = $Workbook.Sheets.Item(1).Cells.Item($i,$x).Text
                if([String]::IsNullOrEmpty($value)) {
                    break
                }
                if($allvalueshash.ContainsKey($_)) {
                    $allvalueshash[$_].Add($value) | Out-Null
                }
                else {
                    $allvalueshash[$_] = New-Object System.Collections.ArrayList
                }
            }
        }
    }catch{
        Write-Host "Unable to read Excel! Error: $($error.exception.message)" -foregroundcolor red
    }
}

function Prepare-Creation {
    Prepare-Excel
    if($CompanyName.Count -eq 0) {
        if((Read-Host "If you want to automatically populate CompanyName, please enter 'yes'") -eq "yes") {
            $fillinghash.Add("CompanyName",$allvalueshash["CompanyName"])
        }
    }
    else {
        $fillingAttributes.Add("CompanyName",$CompanyName)
    }

    if($Department.Count -eq 0) {
        if((Read-Host "If you want to automatically populate Department, please enter 'yes'") -eq "yes") {
            $fillinghash.Add("Department",$allvalueshash["Department"])
        }
    }
    else {
        $fillingAttributes.Add("Department",$Department)
    }

    if($JobTitle.Count -eq 0) {
        if((Read-Host "If you want to automatically populate JobTitle, please enter 'yes'") -eq "yes") {
            $fillinghash.Add("JobTitle",$allvalueshash["JobTitle"])
        }
    }
    else {
        $fillingAttributes.Add("JobTitle",$JobTitle)
    }

    if($GivenName.Count -eq 0) {
        $fillinghash.Add("GivenName",$allvalueshash["GivenName"])
    }
    else {
        $fillingAttributes.Add("GivenName",$GivenName)
    }

    if($Surname.Count -eq 0) {
        $fillinghash.Add("Surname",$allvalueshash["Surname"])
    }
    else {
        $fillingAttributes.Add("Surname",$Surname)
    }
}

function Create-Users {
    try{
        for($counter = 0;$counter -lt $usercount; $counter++) {
            $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.Password = "Pao9,01!4lsfrd"
            $exprop = New-Object "System.Collections.Generic.Dictionary``2[System.String,System.String]"
            foreach($attribute in $fillinghash.Keys) {
                $values = $fillinghash[$attribute]
                $max = $values.Count-1 
                $random = Get-Random -Minimum 0 -Maximum $max
                $finalvalue = $values[$random]
                $exprop.Add($attribute,$finalvalue)
            }
            $username = $exprop["GivenName"] +'.'+ $exprop["Surname"]
            $upn = $username + '@' + $Domain
            New-AzureADUser -DisplayName $exprop["GivenName"] -PasswordProfile $PasswordProfile -UserPrincipalName $upn -AccountEnabled $true -MailNickName $username -ExtensionProperty $exprop
        }
    }catch{
        Write-Host "Could not create a new AD User! Error: $($error.exception.message)" -foregroundcolor red
    }
    
}

function main {
    Prepare-Execution
    Connect-Services
    Prepare-Creation
    Create-Users
}

main