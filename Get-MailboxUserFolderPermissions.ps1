<#
.SYNOPSIS
Get-MailboxUserFolderPermissions.ps1
.DESCRIPTION 
Retrieves permissions to all user relevant folders of a mailbox,
can be narrowed down to the permissions of a single user
REQUIRES EXCHANGE MANAGEMENT CONSOLE!
.OUTPUTS
Console output like Get-MailboxFolderPermission
.PARAMETER Mailbox
The mailbox that the folder permissions will be retreived from.
SamAccountName recommended.
.PARAMETER User
The user whose permissions to the other user's mailbox folders you want to know.
Must be a mailbox user, 'Default' or 'Anonymous'.
SamAccountName recommended.
.EXAMPLE
.\Get-MailboxUserFolderPermissions.ps1 -Mailbox jobs -User gates
This will retreive all user folder permissions in mailbox "jobs" which are set for user "gates".
.EXAMPLE
.\Get-MailboxUserFolderPermissions.ps1 -Mailbox jobs
This will retreive all user folder permissions in mailbox "jobs"
.NOTES
Written by: Maximilian Otter, 2018
Based on Add-MailboxFolderPermissions written by Paul Cunningham
#>

#requires -version 2

[CmdletBinding()]
param (

    # Parameter MAILBOX
    # Describes the mailbox, to the folders of which the permissions will be set.
    # Use anything that identifies a mailbox, SamAccountName is recommended.
    [Parameter(Mandatory=$true,Position=0)]
    [ValidateScript(
        {        
            if ([bool](Get-Command Get-Mailbox)) {
                if ([bool](Get-Mailbox $_ -ErrorAction SilentlyContinue)) {
                    $true
                } elseif (([bool](Get-Command Get-RemoteMailbox)) -and ([bool](Get-RemoteMailbox $_ -ErrorAction SilentlyContinue))) {
                    Throw "$_ is a cloud mailbox. Please rerun the script from a cloud enabled powershell instance."
                } else {
                    Throw "$_ does not resolve to a mailbox on exchange!"
                }
            } else {
                Throw "Exchange Cmdlets not available!"
            }
        }
    )]
	[string]$Mailbox,
    
    # Parameter USER
    # Describes the user for whom the permissions to the folder of -Mailbox will be set.
    # Use anything that identifies a mailbox, SamAccountName is recommended.
    [Parameter(Mandatory=$false,Position=1)]
    [ValidateScript(
        {
            if ([bool](Get-Command Get-Mailbox)) {
                if (
                    ([bool](Get-Mailbox $_ -ErrorAction SilentlyContinue)) `
                    -or (([bool](Get-Command Get-RemoteMailbox)) -and ([bool](Get-RemoteMailbox $_ -ErrorAction SilentlyContinue))) `
                    -or ([bool](Get-DistributionGroup $_ -ErrorAction SilentlyContinue)) `
                    -or ([bool](Get-MailUser $_ -ErrorAction SilentlyContinue)) `
                    -or ($_ -eq 'Anonymous') `
                    -or ($_ -eq 'Default')
                    ) {
                    $true
                } else {
                    Throw "$_ does not resolve to a mailbox, remote mailbox, distribution group, mailuser, Default or Anonymous!"
                }
            } else {
                Throw "Exchange Cmdlets not available!"
            }
        }
    )]
	[string]$User
)

#...................................
# inline Requirements
#...................................

# Cancel execution if the calling powershell is no Exchange Console
# imported sessions do NOT suffice!
if (!((($exbin) -and ($exbin -clike '*Exchange Server*')) -or ($ConnectionUri -like 'https://outlook.office365.com/PowerShell*'))) {
    Throw 'Script must be run in Exchange Console!'
}


#...................................
# inline Functions
#...................................

# Convert the mailbox's primarysmtpaddress and folderpath
# to the folder identity needed for Get-MailboxFolderPermission
function Get-MailboxFolderIdentity {
param (
    [Parameter(Mandatory=$true)]
    [ValidateScript(
        {
            [bool]([System.Net.Mail.MailAddress]::New("$_"))
        }
    )]
    [string]$PrimarySMTPAddress,

    [Parameter(Mandatory=$true)]
    $FolderInfo
)
<#     if ($FolderInfo.FolderType -cne "Root") {
        $FolderInfo.FolderPath = $FolderInfo.FolderPath -Replace '/','\'
    } else {
        $FolderInfo.FolderPath = '\'
    } #>
    "$($PrimarySMTPAddress):$($FolderInfo.FolderID)"
}


#...................................
# Variables
#...................................

# These are all folder types we will set access to.
# More exist, but are not relevant imho
$IncludedFolderTypes = @(
                        'Calendar',
                        'CommunicatorHistory',
                        'Contacts',
                        'DeletedItems',
                        'Drafts',
                        'Files',
                        'Inbox',
                        'JunkEmail',
                        'Notes',
                        'Outbox',
                        'Root',
                        'SentItems',
                        'Tasks',
                        'User Created'
                        )


$PrimarySMTPAddress = Get-Mailbox $Mailbox `
                        | Select-Object -ExpandProperty PrimarySMTPAddress
if ($PrimarySMTPAddress -isnot [string]) {
    $PrimarySMTPAddress = $PrimarySMTPAddress.Address
}

#...................................
# Script
#...................................

# Get Mailboxfolders: Paths and Types. Types are necessary
# to work independently of the language of the foldernames
Write-Host 'Collecting folder information ...'
$MailboxFolders = Get-MailboxFolderStatistics $Mailbox `
                    | Where-Object {$IncludedFolderTypes -ceq $_.FolderType} `
                    | Select-Object FolderID,FolderPath,FolderType
$MailboxFolderCount = $MailboxFolders.count
# Distinguish, if a $user paramater was supplied or not (different call to Get-MailboxFolderPermission)
$count = 0
Write-Progress -Activity "Enumerating $MailboxFolderCount folders ..." -Status "$count folders processed" -PercentComplete 0 -ID 1
if ([bool]$PSBoundParameters.user) {
    foreach ($MailboxFolder in $MailboxFolders)
    {
        $count++
        Write-Progress -Activity "Enumerating $MailboxFolderCount folders ..." -Status "$count folders processed" -PercentComplete ($count*100/$MailboxFolderCount) -ID 1
        $FolderIdentity = Get-MailboxFolderIdentity -PrimarySMTPAddress $PrimarySMTPAddress -FolderInfo $MailboxFolder
        Get-MailboxFolderPermission -Identity $FolderIdentity -User $User -ErrorAction SilentlyContinue `
        | Select-Object @{Name='FolderPath';Expression={$MailboxFolder.FolderPath}},FolderName,User,AccessRights
    }
} else {
    foreach ($MailboxFolder in $MailboxFolders)
    {
        $count++
        Write-Progress -Activity "Enumerating $MailboxFolderCount folders ..." -Status "$count folders processed" -PercentComplete ($count*100/$MailboxFolderCount) -ID 1
        $FolderIdentity = Get-MailboxFolderIdentity -PrimarySMTPAddress $PrimarySMTPAddress -FolderInfo $MailboxFolder
        Get-MailboxFolderPermission -Identity $FolderIdentity  -ErrorAction SilentlyContinue `
        | Select-Object @{Name='FolderPath';Expression={$MailboxFolder.FolderPath}},FolderName,User,AccessRights
    }
}
Write-Progress -Activity "Enumerating $MailboxFolderCount folders ..." -Completed -ID 1

#...................................
# End
#...................................