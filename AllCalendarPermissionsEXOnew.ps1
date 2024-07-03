#I can't take credit for the bones of the script, though I've updated it for the 'new' EXO- scripts

#Connect to O365 Exchange online using modern authentication
Connect-ExchangeOnline

#Grab a list of all account that might contain calendar folders
$MBList = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object -Property Displayname,Alias,Identity,PrimarySMTPAddress,RecipientTypeDetails
#$MBList = Get-EXOMailbox -identity sjpalm@gih.org -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object -Property Displayname,Alias,Identity,PrimarySMTPAddress,RecipientTypeDetails

$allPermissions = @()
$count = 1; $PercentComplete = 0; $i=1; $j=-1;
foreach ($MB in $MBList) {
    $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
    $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
    $PercentComplete = ($count / @($MBList).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++
    
    #For each mailbox, we grab ALL the calendars contained therein
    $UserCalendars = Get-EXOMailboxFolderStatistics $MB.PrimarySmtpAddress.ToString() -FolderScope Calendar
    foreach ($usercal in $UserCalendars) {
        $j++
        if(($usercal.Name -notlike "*holidays*") -and ($usercal.Name -notlike "*birthday*")){
            if($usercal.Name -eq "Calendar"){
                #This checks for the standard "Calendar" issued to all user mailboxes
                $usercalIdentity = $MB.Alias+":\"+$usercal.Name
                }
            else {
                #This grabs the identity of user-created calendars
                $usercalIdentity = $MB.Alias+":\Calendar\"+$usercal.Name
                }
            #We grab all the permissions associated with sharing the calendar
            $usercalRights = Get-EXOMailboxFolderPermission -Identity $usercalIdentity | ? {$_.User.DisplayName -notin @("Owner@local","Member@local")} #You can adjust this to filter out Default & Anonymous entries
            if (!$usercalRights) { 
                $objPermissions = New-Object PSObject
                $i++
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox Name" -Value $MB.DisplayName
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Calendar folder" -Value $usercal.Name
                    
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value ''
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value ''
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permission Type" -Value ''
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "FolderSize" -Value '$UserCalendars[$j].Foldersize'
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ItemCount" -Value '$UserCalendars[$j].ItemsInFolderAndSubfolders'

                $allPermissions += $objPermissions
            
            }
            foreach ($entry in $usercalRights) {
                #Prepare the output object
                $objPermissions = New-Object PSObject
                $i++
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox Name" -Value $MB.DisplayName
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails.ToString()
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Calendar folder" -Value $usercal.Name
                   
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $entry.User
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value $($entry.AccessRights -join ";")
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "FolderSize" -Value $UserCalendars[$j].Foldersize

                $allPermissions += $objPermissions

            }  
        }
    #Appends all the permissions for a particular calendar to CSV, clears it, and then move on to the next one.
    #Output is written to same directory from which the script is called.
    $allPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_CalendarPermissions2.csv" -Append -NoTypeInformation -Encoding UTF8 -UseCulture
    $allPermissions = @()
    }
    $j=-1
}
