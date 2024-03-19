# Load assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Function to show a pop-up message
function Show-PopupMessage {
    param(
        [string]$Message,
        [string]$Title
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title)
}

# Notify about script execution in 5 minutes
Show-PopupMessage -Message "The script will run in 5 minutes. Please close all necessary applications." -Title "Script Notification"

# Wait for 5 minutes
Start-Sleep -Seconds 3

# Define the directories and Tableau Prep CLI command
$dataSourceZip = "D:\OneDrive - Hanoi University of Science and Technology\document\project\data_source_zip"
$dataSource = "D:\OneDrive - Hanoi University of Science and Technology\document\project\data_source"
$dataWarehouse = "D:\OneDrive - Hanoi University of Science and Technology\document\project\data_warehouse"
$etlFlowPath = "D:\OneDrive - Hanoi University of Science and Technology\document\project\ETL_flow.tfl"
$jsonFlowPath = "D:\OneDrive - Hanoi University of Science and Technology\document\project"
$startingIdFile = "D:\OneDrive - Hanoi University of Science and Technology\document\project\data_source\lookup_data\StartingID.csv"
$tableauPrepCmd = "D:\program file\Tableu Prep Buider 2023\scripts\tableau-prep-cli.bat"

Write-Output "Checking for new ZIP files in data_source_zip folder..."

# Get the list of base names of CSV files in data_source
$existingCsvBaseNames = Get-ChildItem -Path $dataSource -Filter "*.csv" | ForEach-Object { $_.BaseName }

# Check for new ZIP files in data_source_zip folder
$newZipFiles = Get-ChildItem -Path $dataSourceZip -Filter "*.zip" | Where-Object { $_.BaseName -notin $existingCsvBaseNames }

foreach ($zipFile in $newZipFiles) {
    Write-Output "Processing file: $($zipFile.Name)"

    # Unzip the file directly into the data_source folder
    Expand-Archive -Path $zipFile.FullName -DestinationPath $dataSource
    Write-Output "Unzipped file to data_source."

    # Determine the name of the extracted CSV file
    # Assuming there's only one CSV file in each ZIP
    $extractedCsvFile = Get-ChildItem -Path $dataSource -Filter "*.csv" | Where-Object { $_.CreationTime -gt (Get-Date).AddMinutes(-1) } | Select-Object -First 1

    # Rename the extracted CSV file to have the same base name as the ZIP file
    $newCsvFileName = "$dataSource\" + $zipFile.BaseName + ".csv"
    Rename-Item -Path $extractedCsvFile.FullName -NewName $newCsvFileName
    Write-Output "Renamed extracted file to match ZIP file name."

    # Get the starting ID from the CSV file and increment by 1 for each new file
    $startingId = ((Get-Content -Path $startingIdFile | Select-Object -Skip 1 -First 1) -as [int])
    $startingId += 1

    # Delete the StartingID.csv file after reading
    Remove-Item -Path $startingIdFile
    Write-Output "Deleted StartingID.csv file."

    # The extracted CSV file will have the same base name as the ZIP file
    $csvFileName = "$dataSource\" + $zipFile.BaseName + ".csv"
    $namefileParameter = $zipFile.BaseName

    # Prepare parameters for the JSON file
    $parameters = @{
        "filename" = $zipFile.BaseName;
        "startingid" = $startingId
    }

    # Convert parameters to JSON format
    $jsonParameters = $parameters | ConvertTo-Json

    # Define the path for the parameters JSON file
    $jsonFilePath = "$jsonFlowPath\parameters.override.json"

    # Write the JSON content to the file
    $jsonParameters | Out-File -FilePath $jsonFilePath -Force


    # Run Tableau Prep flow with parameters
    Write-Output "Running Tableau Prep flow..."
    & $tableauPrepCmd -t $etlFlowPath -p $jsonFilePath
    Write-Output "Tableau Prep flow completed."

    # Updating Dim_Airport, Dim_Carrier, Dim_Route files
    $finalAirportFile = "$dataWarehouse\Dim_Airport.hyper"
    $tempAirportFile = "$dataWarehouse\Temp_Dim_Airport.hyper"
    Copy-Item -Path $finalAirportFile -Destination $tempAirportFile -Force

    $finalCarrierFile = "$dataWarehouse\Dim_Carrier.hyper"
    $tempCarrierFile = "$dataWarehouse\Temp_Dim_Carrier.hyper"
    Copy-Item -Path $finalCarrierFile -Destination $tempCarrierFile -Force

    $finalRouteFile = "$dataWarehouse\Dim_Route.hyper"
    $tempRouteFile = "$dataWarehouse\Temp_Dim_Route.hyper"
    Copy-Item -Path $finalRouteFile -Destination $tempRouteFile -Force

    # Log or output the successful update
    Write-Output "Processed and updated data for file: $csvFileName"
}

if ($newZipFiles.Count -eq 0) {
    Write-Output "No new ZIP files to process."
}

# Notify when the script has run successfully
Show-PopupMessage -Message "Script execution completed successfully." -Title "Script Completed"
exit $LastExitCode
