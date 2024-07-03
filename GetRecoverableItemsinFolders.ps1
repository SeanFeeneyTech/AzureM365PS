$text = @"
user1
user2
user3
"@

# Split the text into lines and trim whitespace
$usernames = $text -split "`n" | ForEach-Object { $_.Trim() }

# Output CSV file path
$outputCsvPath = "FolderoutputFileNew.csv"

# Initialize an array to store the results
$results = @()

# Initialize the progress counter
$totalUsers = $usernames.Count
$currentProgress = 0


# Iterate through each username
foreach ($username in $usernames) {
    # Update the progress counter
    $currentProgress++
    $percentComplete = ($currentProgress / $totalUsers) * 100

    # Display progress
    Write-Progress -Activity "Processing Users" -Status "Processing $username" -PercentComplete $percentComplete

    # Get RecoverableItems folder statistics for the user
    $folderStatistics = Get-MailboxFolderStatistics -Identity $username -FolderScope RecoverableItems
    # Iterate through each RecoverableItems subfolder
    foreach ($folderStat in $folderStatistics) {
        # Build an object with the required information
        $result = [PSCustomObject]@{
            User                      = $username
            FolderName                = $folderStat.Name
            FolderAndSubfolderSize    = $folderStat.FolderAndSubfolderSize
            ItemsInFolderAndSubfolders = $folderStat.ItemsInFolderAndSubfolders
        }

        # Add the result to the results array
        $results += $result
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation
