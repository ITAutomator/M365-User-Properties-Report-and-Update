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
#$props = "UserPrincipalName","DisplayName","TelephoneNumber","Mobile","JobTitle","CompanyName","StreetAddress","City","State","PostalCode","Country"
$props = @(
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
if (!(Test-Path $scriptCSV))
{
    ######### Template
    $($props -join ",") | Add-Content $scriptCSV
    "jsmith@domain.com,"+$(($props[1..$($props.Count)] | ForEach-Object{"New value or <clear> or leave blank to keep as is"}) -join ",") | Add-Content $scriptCSV
    #$($props -join "New value or <clear> or leave blank to keep as is,") | Add-Content $scriptCSV
    ######### Template
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
# see if there's a more recent CSV with this naming scheme
$scriptCSV = Get-ChildItem -Path $scriptcsv.Replace(".csv","*.csv") | Sort-Object LastWriteTime | Select-Object -Last 1 | Select-Object FullName -ExpandProperty FullName
# ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV -Encoding UTF8)
$entriescount = $entries.count
# gather a list of properties to be inspected
$entries_cols = ($entries | Get-Member | Where-Object -Property "MemberType" -EQ "NoteProperty" | Where-Object -Property "Name" -NE "UserPrincipalName" | Select-Object "Name").Name
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
Write-Host "Possible column names are:"
Write-Host "UserPrincipalName (Required),$($props -join ",")"  -ForegroundColor Green
Write-Host "Columns found:"
Write-Host ($entries_cols -join ", ") -ForegroundColor Green
Write-Host 'Use ""        to leave column as is (no change)'
Write-Host 'Use "<clear>" to clear column of contents'
Write-Host ""
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"
$required_col = "UserPrincipalName"
if (-not ($entries | Get-Member | Where-Object -Property "Name" -EQ $required_col)) {
    Write-Host "Err: Required column not found: $($required_col)"; PressEnterToContinue; Exit
}
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()
# Load required modules
$module= "Microsoft.Graph.Users" ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
if ($lm_result.startswith("ERR")) {
    Write-Host "ERR: Load-Module $($module) failed. Suggestion: Open PowerShell $($PSVersionTable.PSVersion.Major) as admin and run: Install-Module $($module)";PressEnterToContinue; Return $false
}
# Connect
$myscopes=@()
$myscopes+="User.ReadWrite.All"
$myscopes+="GroupMember.ReadWrite.All"
$myscopes+="Group.ReadWrite.All"
$connected_ok = ConnectMgGraph -myscopes $myscopes
$domain_mg = Get-MgDomain -ErrorAction Ignore| Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
if (-not ($connected_ok))
{ # connect failed
    Write-Host "Connection failed."
}
else
{ # connect OK
    $processed=0
    $choiceLoop=0
    $i=0
    $change_i=0
    foreach ($x in $entries)
    { # each entry
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
        { # Process all not selected yet, Ask
            $choices = @("&Yes","Yes to &All","&No","No and E&xit") 
            $choiceLoop = AskforChoice -Message "Process entry $($i)?" -Choices $choices -DefaultChoice 0
        } # Process all not selected yet, Ask
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
        { # Process
            $processed++
		    ####### Start code for object $x
            $UserNameOrEmail = $x.UserPrincipalName
            $user = Get-MgUser -Filter "(UserPrincipalName eq '$($UserNameOrEmail)')" -Property (@("id")+$entries_cols)  # Get-MgUser -Filter "(mail eq '$($UserNameOrEmail)') or (displayname eq '$($UserNameOrEmail)')"
            if (-not $user)
            { # user bad
                Write-Host "User not found: $($x.UserNameOrEmail) ERR"  -ForegroundColor Red
            } # user bad
            else
            { # user ok
			    ####### Display 'before' info
                Write-host "[Before]"
                ($user | Select-Object $entries_cols | Format-List | Out-String) -Split "`r`n" | Where-Object({ $_ -ne "" }) | Write-Host
			    #######
                $change_made = $false
                ForEach ($prop in $entries_cols)
                { # each prop
                    If ($x.$prop -eq "")
                    { #No data
                    }
                    ElseIf ($x.$prop -eq "<clear>")
                    {
                        if (($user.$prop -eq "") -or ($null -eq $user.$prop))
                        {
                            Write-Host "$($prop): $($user.$prop) <clear> [Already OK]"
                        }
                        else
                        {
                            Write-Host "$($prop): $($user.$prop) <clear>" -ForegroundColor Yellow
                            # Can't use Update-MgUser to set to null, use Invoke-MgGraphRequest instead
                            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/{$($User.id)}" -Body @{$($prop) = $null}
                            $change_made = $true
                        }
                    }
                    ElseIf ($x.$prop -eq $user.$prop)
                    {  #No update
                    } 
                    Else
                    { # Update
                        Write-Host "$($prop): [$($user.$prop)] changed to [$($x.$prop)]" -ForegroundColor Yellow
                        $myargs = @{
                          UserId = $user.Id
                          $prop = $x.$prop
                        }
                        Update-MgUser @myargs
                        $change_made = $true
                    } # Update
                } # each prop
                if ($change_made)
                { # change made
                    $change_i+=1
                    ####### Display 'after' info
                    Write-host "[After]"
                    $user = Get-MgUser -Filter "(UserPrincipalName eq '$($UserNameOrEmail)')" -Property (@("id")+$entries_cols) 
                    ($user | Select-Object $entries_cols | Format-List | Out-String) -Split "`r`n" | Where-Object({ $_ -ne "" }) | Write-Host
			        Write-host "[OK]" -NoNewline -ForegroundColor Yellow; Write-host " Change made"
                } # change made
                else
                { # no change
                    Write-host "[OK]" -NoNewline -ForegroundColor Green; Write-host " Nothing changed"
                } # no change
            } # user ok
            ####### End code for object $x
        } # Process
        if ($choiceLoop -eq 2)
            {
            write-host ("Entry "+$i+" skipped.")
            }
        if ($choiceLoop -eq 3)
            {
            write-host "Aborting."
            break
            }
    } # each entry
    Write-Host "------------------------------------------------------------------------------------"
    Write-Host "Changes made: $($change_i)" -ForegroundColor $(if ($change_i -eq 0) {"Green"} else {"Yellow"})
    Write-Host "------------------------------------------------------------------------------------"
    $message ="Done. " +$processed+" of "+$entriescount+" entries processed. Press [Enter] to exit."
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
} # connect OK
PressEnterToContinue