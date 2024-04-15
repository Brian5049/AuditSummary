$initialLocation = Get-Location

# Get the user input passed from Python
$folder = $args[0] + ".csv"

$UserName = $Env:UserName

# Set new location for where the file will be
Set-Location "C:\Users\$UserName\Desktop\AuditSummaries"
$data = Import-Csv $folder -Header Column1,Column2,Column3

# Create a new hash table to store summarized data
$summary = @{}

# Variable to count total devices audited
$totalDevicesAudited = 0

# Loop through the data and summarize
foreach ($row in $data) {
    $device = $row.Column1
    $response = $row.Column2
    $notes = $row.Column3

    # Increment total devices audited count
    $totalDevicesAudited++

    # This was just troubleshooting for my instance. You probably won't
    # need exactly this but will need to debug around it.
    # I left it here for you to see what I was working with.
    if ($summary.ContainsKey($device)) {
        $summary[$device].YesCount += if ($response -eq "Yes") {1} else {0}
        $summary[$device].NoCount += if ($response -eq "No") {1} else {0}
        $summary[$device].NotNullCount += if ($notes -ne "") {1} else {0}
    } else {
        $summary[$device] = @{
            YesCount = if ($response -eq "Yes") {1} else {0}
            NoCount = if ($response -eq "No") {1} else {0}
            NotNullCount = if ($notes -ne "") {1} else {0}
        }
    }
}

$filenameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($folder)

$outputCSVFile = "$filenameWithoutExtension-summary.csv"

# Get the last line of the original CSV
$lastLine = Get-Content $folder | Select-Object -Last 1

# Write the last line of the original CSV as the first line of the new CSV
$lastLine | Out-File -FilePath $outputCSVFile -Append

# These were my headers, yours will most likely be different
"Audit Questions,YesCount,NoCount,NotNullCount" | Out-File -FilePath $outputCSVFile -Append

# Write the summary to the new CSV file
$summary.GetEnumerator() | ForEach-Object {
    $device = $_.Key
    $yesCount = $_.Value.YesCount
    $noCount = $_.Value.NoCount
    $notNullCount = $_.Value.NotNullCount
    # Check if the device name contains a comma, if yes, encapsulate in quotes
    if ($device -match ",") {
        $device = "`"$device`""
    }
    "$device,$yesCount,$noCount,$notNullCount" | Out-File -FilePath $outputCSVFile -Append
}

Write-Host "Summary CSV file created: $outputCSVFile"

# Getting what I needed
$csvContent = Get-Content $outputCSVFile

# Throwing away what I don't
$csvContent = $csvContent | Select-Object -First ($csvContent.Count - 1)

# Excel objects suck... TRASH!
# Again, these objects were specific to my project. You will 
# most likely need to figure out what yours are
$csvContent = $csvContent | Where-Object { $_ -notmatch '^P4' }

# Set everything up for the big show!
$csvContent | Set-Content $outputCSVFile

# Just doing some house cleaning
Remove-Item $folder

# Move back to where it all began...
Set-Location $initialLocation
