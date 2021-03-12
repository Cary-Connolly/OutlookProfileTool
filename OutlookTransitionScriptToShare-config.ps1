#region Global Variables/Settings
$Script:LoggingEnabled = $true
$Script:LogFileName = ($env:USERNAME + '-OutlookTransition.Log')
$Script:MigrationScheduledTaskName = "Outlook 365 Transition - $env:USERNAME"
$Script:OutlookEXEPath = 'C:\Program Files (x86)\Microsoft Office\Office16\Outlook.exe'
$Script:HKCUOutlookKey = 'HKCU:\Software\Microsoft\Office\16.0\Outlook'
$Script:NewOutlookProfileName = 'Default O365 Profile'

$Script:MigrateMsg = "Your new M365 mailbox is ready to be set up.`r`n`r` This process will take a few minutes. `r`n`r`Please close all Outlook windows and click OK to begin. `r`n`r`Votre nouveau service de courriel M365 est prêt à être configuré.`r`n`r` Ce processus prendra quelques minutes. `r`n`r`Veuillez fermer toutes les fenêtres Outlook et cliquer sur OK pour commencer."

$Script:MigrateMsgOutlookOpened = "You need to be on VPN for the process to run.`r`n`r`  You will be prompted to click OK to launch Outlook when the setup is complete. `r`n`r`Vous devez être connecté à un RVP pour que le processus fonctionne.`r`n`r` Vous serez invité à ouvrir Outlook lorsque la configuration sera terminée. "

$Script:MigrateOutlookFailedToClose = "Please completely close Outlook, including any Outlook pop-up windows and click OK. `r`n`r`Veuillez fermer complètement Outlook, y compris toute fenêtre contextuelle Outlook et cliquer sur OK."

$Script:MigrationCompleteMsg = "Your M365 mailbox move has completed! Emails could take a few minutes to download.`r`n`r` Click OK to launch Outlook.  `r`n`r`Note that Outlook log in will be automatic for most, although some may be prompted to log in.`r`n`r`  If so, please use  your new email address and the password you use to log in to your computer.`r`n`r`La migration de votre service de courriel M365 est terminée! Le téléchargement des courriels pourrait prendre quelques minutes.`r`n`r` Cliquez sur OK pour lancer Outlook.`r`n`r` Veuillez noter que le lancement d’Outlook se fera automatiquement pour la plupart, bien que certains peuvent être invité à ouvrir une session.`r`n`r` Si oui, veuillez utiliser votre nouvelle adresse de courriel et le mot de passe que vous utilisez pour vous connecter à votre ordinateur."

$Script:MigrationCompleteMsgNoNetwork = "Your M365 mailbox move has completed! Please connect to your secure remote access (e.g. VPN), then open Outlook.`r`n`r` Emails could take a few minutes to download.`r`n`r`  If prompted to log in to Outlook, please use  your new email address and the password you use to log in to your computer.`r`n`r`La migration de votre service de courriel M365 est terminée! Veuillez vous connecter à votre accès à distance sécurisé (par ex. RPV), puis ouvrir Outlook.`r`n`r` Le téléchargement des courriels pourrait prendre quelques minutes.  `r`n`r`Si vous êtes invité à ouvrir une session, veuillez utiliser votre nouvelle adresse de courriel et le mot de passe que vous utilisez pour vous connecter à votre ordinateur."

#endregion Global Variables/Settings
#region Functions

$XMLfile = ".\config.xml"
[xml]$Deps=Get-content $XMLfile

$Script:LoggingFolderPath = $DEps.Depts.Department.LoggingFolderPath
$Script:MigrationADGroup = $DEps.Depts.Department.MigrationADGroup
$Script:InternalAddressToPing = $DEps.Depts.Department.InternalAddressToPing
$Script:MsgTitle = $DEps.Depts.Department.MsgTitle

# Dialog Box
#$wsh = New-Object -ComObject Wscript.Shell
#$wsh.Popup($MigrateMsg)
#$wsh.Popup($MigrateMsgOutlookOpened)
#$wsh.Popup($MigrateOutlookFailedToClose)
#$wsh.Popup($MigrationCompleteMsg)
#$wsh.Popup($MigrationCompleteMsgNoNetwork)

# Log file
function Test-LogFolder {
	If ($Script:LoggingEnabled -eq $true) {
		If (-not (Test-Path $Script:LoggingFolderPath -ErrorAction SilentlyContinue)) {
			New-Item -ItemType Directory -Path $Script:LoggingFolderPath | Out-Null
			If (-not (Test-Path $Script:LoggingFolderPath -ErrorAction SilentlyContinue)) {
				$Script:LoggingEnabled = $false
				return
			}
		}
		$Script:LoggingEnabled = $true
		return
    }
}
function Write-Log ($entry) {
	If ($Script:LoggingEnabled -eq $true) {
		((Get-Date).ToString() + " - $entry") >> "$Script:LoggingFolderPath\$Script:LogFileName"
	}
}

# Prechecks
function Test-AlreadyMigrated {
	Write-Log "Checking existance of ZeroConfigExchangeOnce at $Script:HKCUOutlookKey\AutoDiscover"
	Try {
		If ((Get-ItemPropertyValue -Path "$Script:HKCUOutlookKey\AutoDiscover" -Name ZeroConfigExchangeOnce -ErrorAction SilentlyContinue) -as [bool]) {
			Write-Log ($env:USERNAME + ' has already migrated')
			return $true
		} Else {
			Write-Log ($env:USERNAME + ' has not migrated yet')
			return $false
		}
	}
	Catch {
		Write-Log ($env:USERNAME + ' has not migrated yet')
		return $false
	}
}
function Test-OutlookEXE {
	If ((Test-Path $Script:OutlookEXEPath) -as [bool]) {
		Write-Log "$Script:OutlookEXEPath exists"
		return $true
	} Else {
		Write-Log "$Script:OutlookEXEPath does not exist"
		return $false
	}
}
function Test-OutlookProfileExists {
	Write-Log "Checking existance of Outlook profiles at $Script:HKCUOutlookKey\Profiles"
	If ((get-childitem -path "$Script:HKCUOutlookKey\profiles" -ErrorAction SilentlyContinue) -as [bool]) {
		Write-Log ('Found Outlook profile(s) for ' + $env:USERNAME)
		return $true
	} Else {
		Write-Log ('No Outlook profile(s) found for ' + $env:USERNAME)
		return $false
	}
}
function Test-InternalNetwork {
	If ((Test-Connection $Script:InternalAddressToPing -Count 1) -as [bool]) {
		Write-Log "Able to ping $Script:InternalAddressToPing"
		return $true
	} Else {
		Write-Log "Not able to ping $Script:InternalAddressToPing"
		return $false
	}
}
function Test-MigrationADGroup {
	$Groups = (New-Object System.DirectoryServices.DirectorySearcher("(&(objectCategory=User)(samAccountName=$($env:username)))")).FindOne().GetDirectoryEntry().memberOf
	If (($Groups -contains $Script:MigrationADGroup) -as [bool]) {
		Write-Log ($env:username + ' is part of AD group ' + $Script:MigrationADGroup)
		return $true
	} Else {
		Write-Log ($env:username + ' is not part of AD group ' + $Script:MigrationADGroup)
		return $false
	}
}

# Cleanup
function Remove-OutlookRegProfiles {
	If (Test-Path $Script:HKCUOutlookKey) {
		Write-Log 'Removing Outlook registry profiles'
		If (Test-Path "$Script:HKCUOutlookKey\Profiles\*") {
			Get-ChildItem -Path "$Script:HKCUOutlookKey\Profiles\*" | Remove-Item -Recurse
			Write-Log 'Successfully removed Outlook registry profiles'
		}
		Else {
			Write-Log 'No Outlook registry profiles to delete'
		}
	}
}
function Remove-OutlookAppdataFiles {
	$OutlookAppdata = $env:LOCALAPPDATA + '\Microsoft\Outlook\'
	Write-Log "Deleting AppData files at $OutlookAppdata"
	If (Test-Path $OutlookAppdata) {
		Get-ChildItem $OutlookAppdata -File -Force | Remove-Item -Force
		If (-not (Get-ChildItem $OutlookAppdata -File -Force)) {
			Write-Log 'AppData files successfully deleted'
		} Else {
			Write-Log 'Unable to delete all AppData files'
		}
	}
}
function Remove-OutlookCredentials {
	Write-Log 'Removing saved Outlook credentials'
	[array]$OutlookCredentials = cmdkey /list | Where-Object { $_ -like "*Target:*" -and $_ -like "*Outlook*" }
	If ($OutlookCredentials) {
		foreach ($Credential in $OutlookCredentials) {
			Write-Log "Removing stored creds: $Credential"
			cmdkey /del:($Credential -replace " ", "" -replace "Target:", "")
		}
	} Else {
		Write-Log 'No saved Outlook credentials to remove'
	}
}

# New profile setup (registry edits)
function Set-NewOffice365Profile {
	Try {
		Write-Log "Creating registry key $Script:NewOutlookProfileName at '$Script:HKCUOutlookKey\Profiles\'"
		new-item -path "$Script:HKCUOutlookKey\Profiles\" -Name $Script:NewOutlookProfileName -Force -ErrorAction Stop | out-Null
		Write-Log "$Script:NewOutlookProfileName successfully created"
	} Catch {
		Write-Log "Error creating $Script:NewOutlookProfileName"
	}
	
	Try {
		Write-Log "Setting registry string 'DefaultProfile' at '$Script:HKCUOutlookKey' to $Script:NewOutlookProfileName"
		New-ItemProperty -Path $Script:HKCUOutlookKey -Name DefaultProfile -Value $Script:NewOutlookProfileName -PropertyType String -Force -ErrorAction Stop | out-Null
		Write-Log "'DefaultProfile' successfully set to $Script:NewOutlookProfileName"
	}
	Catch {
		Write-Log "Error setting 'DefaultProfile' to $Script:NewOutlookProfileName"
	}
	
	Try {
		Write-Log "Creating registry DWord 'ZeroConfigExchangeOnce' at '$Script:HKCUOutlookKey\AutoDiscover' with a value of 1"
		New-ItemProperty -Path "$Script:HKCUOutlookKey\AutoDiscover" -Name ZeroConfigExchangeOnce -Value 1 -PropertyType DWord -Force -ErrorAction Stop | out-Null
		Write-Log "'ZeroConfigExchangeOnce' successfully created"
	}
	Catch {
		Write-Log "Error creating 'ZeroConfigExchangeOnce'"
	}
}

# Message prompt logic
function Show-MigrationPrompt {
	If (Get-Process Outlook -ErrorAction SilentlyContinue) {
		Write-Log 'Outlook.exe running. Prompting user to close Outlook'
		$Result = [Microsoft.VisualBasic.Interaction]::MsgBox($Script:MigrateMsgOutlookOpened, 'OKCancel,SystemModal,Information', $Script:MsgTitle)
		If ($Result -eq 'OK') {
			Write-Log 'User accepted prompt'
			If (Stop-OutlookEXE) {
				return $true
			}
			Else {
				Write-Log 'Informing user that Outlook.exe failed to close'
				$null = [Microsoft.VisualBasic.Interaction]::MsgBox($Script:MigrateOutlookFailedToClose, 'OKOnly,SystemModal,Exclamation', $Script:MsgTitle)
				Write-Log 'User closed prompt'
				Write-Log 'Exit reasoning: Outlook still opened'
				exit
			}
		}
		Else {
			Write-Log 'User cancelled prompt'
			return $false
		}
	}
	Else {
		Write-Log 'Outlook.exe currently not running'
		Write-Log 'Prompting user to migrate'
		$Result = [Microsoft.VisualBasic.Interaction]::MsgBox($Script:MigrateMsg, 'OKCancel,SystemModal,Information', $Script:MsgTitle)
		If ($Result -eq 'OK') {
			Stop-OutlookEXE
			Write-Log 'User accepted prompt'
			return $true
		}
		Else {
			Write-Log 'User cancelled prompt'
			return $false
		}
	}
}
function Show-MigrationComplete {
	Write-Log "Showing 'Migration Complete' MessageBox"
	$null = [Microsoft.VisualBasic.Interaction]::MsgBox($Script:MigrationCompleteMsg, 'OKOnly,SystemModal,Information', $Script:MsgTitle)
	Write-Log "'Migration Complete' MessageBox closed"
}
function Show-NoInternalNetwork {
	Write-Log "Showing 'No VAC Network' MessageBox"
	$null = [Microsoft.VisualBasic.Interaction]::MsgBox($Script:MigrationCompleteMsgNoNetwork, 'OKOnly,SystemModal,Information', $Script:MsgTitle)
	Write-Log "'Migration Complete' MessageBox closed"
	Write-Log "Launching Outlook"
}

# Misc
function Stop-OutlookEXE {
	If ($p = Get-Process Outlook -ErrorAction SilentlyContinue) {
		Write-Log 'Attempting to close Outlook.exe'
		[System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic") | Out-Null
		For ($i = 0; $i -lt 10; $i++) {
			If (-not ($p.HasExited)) {
				Try {
					[Microsoft.VisualBasic.Interaction]::AppActivate($p.ID)
					Start-Sleep -Milliseconds 50
					Write-Log 'Sending "ESC" and "n" key presses to Outlook window'
					[System.Windows.Forms.SendKeys]::SendWait("{ESC}")
					Start-Sleep -Milliseconds 50
					[System.Windows.Forms.SendKeys]::SendWait("n")
					Start-Sleep -Milliseconds 400
				}
				Catch {
					Write-Log 'Unable to focus Outlook.exe window. Might of closed'
					break
				}
			}
			Else {
				Write-Log 'Outlook.exe stopped successfully'
				return $true
			}
		}
		If ($p = Get-Process Outlook -ErrorAction SilentlyContinue) {
			Write-Log 'Outlook still running'
			Write-Log 'Attempting to force close Outlook.exe'
			$p | Stop-Process -Force
			Start-Sleep -Seconds 2
			If (Get-Process Outlook -ErrorAction SilentlyContinue) {
				Write-Log 'Outlook.exe failed to stop'
				return $false
			}
			Else {
				Write-Log 'Outlook.exe stopped successfully'
				return $true
			}
		}
		Else {
			Write-Log 'Outlook.exe stopped successfully'
			return $true
		}
	}
	Else {
		Write-Log 'Outlook.exe currently not running'
		return $true
	}
}
function Disable-MigrationScheduledTask {
	Write-Log "Attempting to disable scheduled task '$Script:MigrationScheduledTaskName'"
	Try {
		Get-ScheduledTask -TaskName $Script:MigrationScheduledTaskName -ErrorAction Stop | Disable-ScheduledTask -ErrorAction Stop | Out-Null
		Write-Log "Scheduled task '$Script:MigrationScheduledTaskName' disabled"
	} Catch {
		Write-Log "Scheduled task '$Script:MigrationScheduledTaskName' doesn't exist or failed to disable"
	}
}

#endregion Functions
#region Main
Add-Type -Path '.\Microsoft.VisualBasic.dll
Add-Type -Path '.\System.Management.Automation.dll
# Test/Setup log folder
Test-LogFolder

Write-Log ' '
Write-Log ('Starting Outlook Transiton Script for user ' + $env:USERNAME)

# Prechecks
If (Test-AlreadyMigrated) { Disable-MigrationScheduledTask; Write-Log 'Exit reasoning: Already migrated'; exit }
If (-not (Test-OutlookEXE)) { Disable-MigrationScheduledTask; Write-Log 'Exit reasoning: Outlook not installed on machine'; exit }
If (-not (Test-OutlookProfileExists)) { Disable-MigrationScheduledTask; Write-Log 'Exit reasoning: No existing Outlook profiles to replace'; exit }
If (-not (Test-InternalNetwork)) { Write-Log 'Exit reasoning: No internal Network'; exit }
If (-not (Test-MigrationADGroup)) { Write-Log 'Exit reasoning: Not in migration group'; exit }

# Prompt user to Migrate
If (-not (Show-MigrationPrompt)) { Write-Log 'Exit reasoning: User cancelled prompt'; exit }

# Cleanup #1
Remove-OutlookRegProfiles

# Profile setup
Set-NewOffice365Profile

# Cleanup #2
Remove-OutlookAppdataFiles
Remove-OutlookCredentials

# Disable scheduled task
Disable-MigrationScheduledTask

# Prompt Migration Complete
Show-MigrationComplete

# Open Outlook (while on network)
If (Test-InternalNetwork) {
	Write-Log "Launching Outlook"
	Start-Process Outlook
}
Else {
	Show-NoInternalNetwork
}

Write-Log 'Exit reasoning: Script Completed'
#endregion Main
