# Start point to return to, get EnvPath, and move to the correct directory
$initialLocation = Get-Location
$UserName = $Env:UserName

# Change this file path to the one you will need to work in. Typically the directory that holds the 
# directories of the files.
$PathToFile = "FilePathFromUserDirectory"

Set-Location "C:\Users\$UserName\$PathToFile"

# Get the user input passed from Python
$folder = $args[0]

# Move into the correct folder to summarize
try {
    Set-Location $folder
} catch {
    Write-Host "Error: $_"
    Exit
}

# Define the folder containing the Excel files
$folderPath = Get-Location

Write-Host "Beginning Loop. Please be patient"

Import-Module ImportExcel

# Initialize variables to store counts
$totalFilesAudited = 0
$countYesP6 = 0
$countNoP6 = 0
$totalNotNullP8 = 0
$results = @()

# Iterate through each Excel file
Get-ChildItem $folderPath -Filter *.xlsx | ForEach-Object {
    $totalFilesAudited++
    $excelFile = $_.FullName
    Write-Host "Processing file: $($_.Name)"
    
    # Edit this variable to tell the script which sheet to read
    $worksheetName = "Pick a Worksheet to pull from"
    
    # Import data from the specific worksheet in the file
    try {
        $data = Import-Excel -Path $excelFile -WorksheetName $worksheetName -NoHeader
    } catch {
        Write-Host "Failed to import workbook '$excelFile' with worksheet '$worksheetName': $_"
        continue
    }

    # Loop through each row in the data:
    # These are specific cell locations to read from
    # You will need to edit to make it read your files
    # correctly. Change the 16, and 32 to reflect your needs
    for ($row = 16; $row -le 32; $row++) {
        $rowData = $data[$row]
        if ($rowData -eq $null) {
            continue
        }
        
        # Extract specific values from the Excel row
        # These are specific to my use case and may not 
        # be relevant to yours. Take care in editing...
        $valueP4 = $rowData.P4
        $countYesP6 += if ($rowData.P6 -eq "Yes") { 1 } else { 0 }
        $countNoP6 += if ($rowData.P6 -eq "No") { 1 } else { 0 }
        
        # This line proves to be extremely important going forward with this script.
        $totalNotNullP8 += if ($rowData.P8 -ne $null) { 1 } else { 0 }
        
        # Add row data to the results array
        $results += [PSCustomObject]@{
            P4 = $valueP4
            P6 = $rowData.P6
            P8 = $rowData.P8
        }
    }
}

# Add the total number of files audited to CSV
# This number will be important later
$results += [PSCustomObject]@{
    P4 = "Total devices audited"
    P6 = $totalFilesAudited
    P8 = ""
}

# Output the total files audited
Write-Host " "
Write-Host "Total devices audited: $totalFilesAudited"
Write-Host " "

# Define the output CSV file path
$outputCSV = "C:\Users\$UserName\Desktop\AuditSummaries\$folder.csv"

# Export the results to a CSV file
$results | Export-Csv -Path $outputCSV -NoTypeInformation

# Move back to where you started
Set-Location $initialLocation
