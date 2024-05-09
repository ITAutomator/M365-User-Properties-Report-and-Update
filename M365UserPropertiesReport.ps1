#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####
#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open
### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
############
####
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in O365"
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf)"
Write-Host "-----------------------------------------------------------------------------"
$no_errors = $true
$error_txt = ""
$results = @()
# Load required modules
$module= "Microsoft.Graph.Users" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
if ($lm_result.startswith("ERR")) {
    Write-Host "ERR: Load-Module $($module) failed. Suggestion: Open PowerShell $($PSVersionTable.PSVersion.Major) as admin and run: Install-Module $($module)";Start-sleep  3; Return $false
}
# Connect
$connected_ok = ConnectMgGraph 
if (!($connected_ok)) 
{ # connect failed
    Write-Host "[connection failed]"
}
else
{ # connect ok
    Write-Host "CONNECTED"
    Write-Host "--------------------"
    $mg_properties = @(
        'id'
        ,'UserPrincipalName'
        ,'DisplayName'
        ,'mail'
        ,'AccountEnabled'
        ,'BusinessPhones'
        ,'city'
        ,'country'
        ,'CreatedDateTime'
        ,'department'
        ,'GivenName'
        ,'JobTitle'
        ,'LastPasswordChangeDateTime'
        ,'MobilePhone'
        ,'OfficeLocation'
        ,'postalcode'
        ,'state'
        ,'streetAddress'
        ,'Surname'
        ,'userType'
    )
    ####
    Write-host "Exporting n properties: $($mg_properties.Count)"
    If (AskForChoice "Include Group membership info? (takes a bit longer)")
    {
        $mg_properties += "Groups"
    }
    $mg_properties += "Manager"
	####### Retrieve Azure AD User list
    $mgusers = Get-MGuser -All -Property $mg_properties
    Write-Host "User Count: $($mgusers.count) [All users]"
    $mgusers = $mgusers | Where-Object UserType -EQ Member
    Write-Host "User Count: $($mgusers.count) [UserType=Members (vs Guests)]"
    $mgusers = $mgusers | Where-Object AccountEnabled -eq $true
    Write-Host "User Count: $($mgusers.count) [AccountEnabled=True]"
    #$mgusers = $mgusers | where-object { $_.LicenseDetails.count -ne 0}
    #Write-Host "User Count: $($mgusers.count) [LicenseDetails]"
    $mgusers = $mgusers | Sort-Object DisplayName
    #### clear rows
    $rows = @()
    $i=0
    $icount=$mgusers.count
    ForEach ($mguser in $mgusers)
    {
        $i++
        Write-host "$($i) of $($icount): $($mguser.DisplayName) <$($mguser.Mail)>"
        # create a new empty row
        $row = New-Object -TypeName psobject
        # append each column needed
        ForEach ($prop in $mg_properties)
        {
            if ($prop -eq "BusinessPhones"){
                $propname = "BusinessPhones"
                $propvalue=($mguser.BusinessPhones|Sort-Object|ForEach-Object {"$($_)"}) -Join ", "
                $row | Add-Member -Type NoteProperty -Name $propname -Value $propvalue
            }
            elseif ($prop -eq "Groups"){
                $propname = "Groups"
                $UserGroups = GroupParents $mguser.Id
                $propvalue = ($UserGroups.displayName | Sort-Object) -join ", "
                $row | Add-Member -Type NoteProperty -Name $propname -Value $propvalue
            }
            elseif ($prop -eq "Manager"){
                $propname = "Manager"
                $mgr_id = Get-MgUserManager -UserId $mguser.Id -ErrorAction Ignore
                if ($mgr_id) {
                    $mgr = $MgUsers | Where-Object Id -EQ $mgr_id.Id
                    $propvalue = "$($mgr.DisplayName) <$($mgr.Mail)>"
                }
                else {
                    $propvalue = ""
                }
                $row | Add-Member -Type NoteProperty -Name $propname -Value $propvalue
            }
            else{
                $row | Add-Member -Type NoteProperty -Name $prop -Value $mguser.($prop)
            }
        }
        ### append row
        $rows+= $row
    }
    ###
    Write-Host "User Count: $($mgusers.count) [AssignedLicenses=True]"
    Write-host "Exporting info to CSV..."
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    $scriptCSVdated= $scriptCSV.Replace(".csv"," $($date).csv")
    if ($PSVersionTable.PSVersion.Major -lt 7)
    { # ps 5 (Excel likes UTF8-Bom CSVs, PS5 defaults utf8 to BOM)
        $rows | Export-Csv $scriptCSVdated -NoTypeInformation -Encoding utf8
    }
    else
    { # ps 7 (Excel likes UTF8-Bom CSVs, PS7 changed utf8 to be NOBOM, so use utf8BOM)
        $rows | Export-Csv $scriptCSVdated -NoTypeInformation -Encoding utf8BOM
    }
	#######
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    Write-host "File: $(split-path $scriptCSVdated -Leaf)" -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------------------------"
    $message ="Done. Press [Enter] to exit."
    Write-Host $message
    Write-Host "------------------------------------------------------------------------------------"
	#################### Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    #################### Transcript Save
} # connect ok
PressEnterToContinue