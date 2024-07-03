Connect-ExchangeOnline
Connect-AzureAD
# Function to display a progress bar
function Write-ProgressBar {
    param(
        [int]$Completed,
        [int]$Total,
        [string]$Activity
    )

    $percentComplete = $Completed / $Total * 100
    $progressBarLength = 100  # Length of the progress bar (10 times longer)
    $progressBar = '=' * [int](($percentComplete / 100) * $progressBarLength)
    $remainingBar = '-' * ($progressBarLength - $progressBar.Length)
    Write-Host -NoNewline -ForegroundColor Green '['
    Write-Host -NoNewline -ForegroundColor Green $progressBar
    Write-Host -NoNewline -ForegroundColor Gray $remainingBar
    Write-Host -NoNewline -ForegroundColor Green ']'
    Write-Host -NoNewline -ForegroundColor Cyan " ${percentComplete:F1}% $Activity"
    Write-Host  # Add a new line
}

# Retrieve the list of Microsoft 365 groups
$groups = Get-AzureADMSGroup -All $true

$groupCount = 1
$totalGroups = $groups.Count

$memberDetailsCSV = "CapLink_AZGroupDetails.csv"

# Process each group
foreach ($group in $groups) {
    Write-ProgressBar $groupCount $totalGroups "Processing group $($group.DisplayName)"

    # Retrieve members and owners, tracking link types
    $members = @(@(Get-AzureADGroupMember -ObjectId $group.Id) | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name LinkType -Value "Member" -PassThru
    })
    $owners = @(@(Get-AzureADGroupOwner -ObjectId $group.Id) | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name LinkType -Value "Owner" -PassThru
    })
    # Process members and owners together
    foreach ($memberOrOwner in $owners + $members) {
            #$MemberOrOwnerUPN = $memberOrOwner.PrimarySmtpAddress -replace "'", "''"
            $MemberOrOwnerUPN = $memberOrOwner.UserPrincipalName
            #$MemberEmail = if ($memberOrOwner.RecipientTypeDetails -eq "GuestMailUser"){($memberOrOwner.ExternalEmailAddress -replace '^SMTP:', '')} else{$memberOrOwner.PrimarySmtpAddress}
            $MemberEmail = $memberOrOwner.mail
            $isTenantGuest = $MemberEmail -match '#EXT#@caplinkorg.onmicrosoft.com'
            if ($isTenantGuest){$MemberType='TenantGuest'} else {$MemberType = $memberOrOwner.RecipientTypeDetails}
            $memberDetails = [PSCustomObject]@{
            Group = $group.DisplayName
            Identity = $memberOrOwner.Identity
            Name = $memberOrOwner.Alias
            MemberType = $MemberType
            MemberEmail = $memberEmail
            Role = $memberOrOwner.LinkType
            Active = if ($memberOrOwner.RecipientTypeDetails -eq "GuestMailUser") {
             "N/A"  # Guest user
            } elseif ($isTenantGuest) {
                  "N/A"  # External Tenant Guest
             } elseif ($memberOrOwner.RecipientTypeDetails -eq "DisabledUser") {
                  $false  # Disabled user
             } elseif ($memberOrOwner.RecipientTypeDetails -eq "SharedMailbox") {
                  "N/A"  # Shared Mailbox - not a user
              } elseif ($memberOrOwner.RecipientTypeDetails -eq "User") {
                  $false  # User not disbled, but no Exchange License user
            } else {(Get-AzureADUser -Filter "UserPrincipalName eq '$memberOrOwnerUPN'").AccountEnabled}
        }
        $memberDetails | Export-Csv $memberDetailsCSV -Append -NoTypeInformation
    }

    $groupCount++
}
