#this is still a bit of a mess. work in progress - use at your own risk; teams me with questions
#to do:
# - (DONE) fix OU move to disabled - need to strip prefix from output during parsing - this is currently commented out
# - (DONE) prepend disabled <date> - to description in local AD
# - (DONE) remove local AD groups
# - (DONE) fix 365 group export
# - (DONE) pull and auto export shared mailboxes Get-Mailbox | Get-MailboxPermission -User <username> | Format-Table -AutoSize *
# - (DONE) export shared mailbox access
# - (DONE) remove manager from org details - check direct reports
# - (DONE) fix disabling OWA - breaks delegate access if done from here ** needed to disable for devices, not completely

# - (DONE/needs attention) clean up exports - generate full report of shared mb, groups, local ad, etc.; include full log of offboarding and convert to something I can mail out 
#           - export BU, Supervisor, departure date, office location, phone number
# - release microsoft licenses
# - have out of office reply use system variables (full name, title, e-mail address)
# - remove user from 365 groups
# - remove shared mailbox access - full and send-as
# - disable office activations
# - remove mobile devices from OWA
# - convert dots to dashes in user name for SPO assignment 
# - auto assign mailnickname if null (this is required to hide from GAL)


Clear-Host

#get creds - user admin acount
#variables that you need to set
$offboardee = $(Write-Host "Who are you offboarding? (username): " -ForegroundColor green -NoNewLine; Read-Host) # Who are you offboarding?

$supervisor = $(Write-Host "Who is their supervisor? (username): " -ForegroundColor green -NoNewLine; Read-Host) # who are you delegating access to?
$orgName = "organization" # orgname prefixes .onmicrosoft.com
$outofoffice ="For questions, please contact xxxxxx"
$outfile = 'C:\scripts\offboarding\'+$offboardee+'_offboarding_report.txt'




#connect to services

#Azure Active Directory
Connect-AzureAD -Credential $credential

#SharePoint Online
#Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -credential $credential

#Exchange Online
#Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -Credential $credential -ShowProgress $true

#Security & Compliance Center
#Connect-IPPSSession -UserPrincipalName $acctName

#Teams and Skype for Business Online *** do we need to remove direct routing?
#Import-Module MicrosoftTeams
#Connect-MicrosoftTeams -Credential $credential



#begin local AD offboarding -----------------------------------------------------------------

#disable local AD account
Disable-ADAccount -Identity $offboardee

#change password
Set-ADAccountPassword -Identity $offboardee -NewPassword (ConvertTo-SecureString -AsPlainText "TempPass3217!" -Force)

#update user description in AD to include disabled date
$current_description = Get-AdUser -Identity $offboardee -Properties Description | Select-Object -ExpandProperty Description
$new_description = "disabled "+(Get-Date -format "yyyy-MM-dd - ")+$current_description
Set-ADUser $offboardee -Description $new_description

#clear manager field in AD
Set-ADUser -Identity $offboardee -Clear manager

#export local security groups
"Local Security Group Membership for: "+$offboardee | Out-File $outfile -Append
Get-ADPrincipalGroupMembership -Identity $offboardee | Out-File $outfile -Append


#remove user from local AD groups
$ADgroups = Get-ADPrincipalGroupMembership -Identity $offboardee | Where-Object {$_.Name -ne "Domain Users"}
if ($null -ne $ADgroups){
	Remove-ADPrincipalGroupMembership -Identity $offboardee -MemberOf $ADgroups -Confirm:$false
}

#hide from GAL
Set-ADUser $offboardee -Add @{msExchHideFromAddressLists="TRUE"}

#Move user to disabled users group-
Get-ADUser -Identity $offboardee | Move-ADObject -TargetPath "OU=Disabled - move after 30 days,OU=AAOUSERS,DC=organization,DC=org"


#---------------- LOCAL AD DONE #



#---------------- BEGIN AZURE/365 #

# Force an Azure AD Sync
Start-ADSyncSyncCycle -PolicyType Delta

# Confirm user sync completed - this just looks to see if the account is disabled in Azure and continues when it can confirm that it is.
Do {
    "Waiting on AzureAD Sync to complete..."
    Start-Sleep -s 5
    $msoluser = (Get-AzureADUser -ObjectId $offboardee'@organization.org').AccountEnabled
   } 
While ($msoluser -eq $true)


# Export shared mailbox membership

"Shared Mailbox Membership for: "+$offboardee | Out-File $outfile -Append
#Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-MailboxPermission |Select-Object Identity,User,AccessRights | Where-Object {($_.user -like $offboardee+'@organization.org')} | Out-File $outfile -Append

$shared_mb_membership = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-MailboxPermission |Select-Object Identity,User,AccessRights | Where-Object {($_.user -like $offboardee+'@organization.org')}
if ($null -ne $shared_mb_membership){
	$shared_mb_membership | Out-File $outfile -Append
}
else {
"No Shared Mailbox Access" | Out-File $outfile -Append
}


# export 365 group membership
"365 Group Membership for: "+$offboardee | Out-File $outfile -Append
$365_membership = get-azureadusermembership -ObjectId $offboardee'@org.org' | Select-Object DisplayName
if ($null -ne $365_membership){
	$365_membership | Out-File $outfile -Append
}
else {
"No 365 Group Membership" | Out-File $outfile -Append
}


# Get info for departing user
$upn        = $offboardee+'@organization.org'

#generate onedrive link and delegate permissions to supervisor
Set-SPOUser -Site https://$orgName-my.sharepoint.com/personal/$offboardee'_organization_org' -LoginName $supervisor'@organization.org' -IsSiteCollectionAdmin $True -ErrorAction SilentlyContinue


#disable activesync, owa, etc
Set-CASMailbox -id $offboardee –ActiveSyncEnabled $false
Set-CASMailbox -id $offboardee –OWAforDevicesEnabled $false

#convert to shared
Set-Mailbox $upn -Type Shared

#delegate mailbox access
Add-MailboxPermission -Identity $upn -User $supervisor'@organization.org' -AccessRights Full -InheritanceType All -AutoMapping $true

#set auto reply
Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled -InternalMessage $outofoffice -ExternalMessage $outofoffice
