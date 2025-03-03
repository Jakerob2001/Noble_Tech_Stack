# Ensure you have the ImportExcel module installed
Import-Module ImportExcel

# Define the property mapping hash table
$propertyMapping = @{
    "ARG" = @("argonauthotel.com")
     "CAB" = @("corazoncabo.com")
     "CHA" = @("chathaminn.com")
    "EDG" = @("edgewaterhotel.com")
    "ELJ" = @("estancialajolla.com")
    "GWC" = @("gatewaycanyons.com")
    "IOF" = @("innonfifth.com")
    "JICR" = @("jekyllclub.com")
    "KKR" = @("sdkonakai.com")
    "LAP" = @("laplayaresort.com")
    "LDM" = @("laubergedelmar.com")
    "LPI" = @("littlepalmisland.com")
    "MAR" = @("marquesa.com")
     "MBR" = @("missionbayresort.com")
    "NHHR" = @("noblehousehotels.com")
    "NVWT" = @("winetrain.com")
    "OKR" = @("oceankey.com")
    "CKA" = @("Pacific City")
	"HCA" = @("Pacific City")
 	"HCL" = @("Pacific City")
     "PBR" = @("pelicanbeach.com")
     "POK" = @("POKEO KEO")
     "POR" = @("hotelportofino.com")
      "RTI" = @("riverterraceinn.com")
      "SOL" = @("solemiami.com")
      "SRC" = @("srsportingclub.com")
      "TR" = @("tetonresorts.com")
      "TSH" = @("thestellahotel.com")
}

foreach ($propertyCode in $propertyMapping.Keys) {
    $newWorkbookPath = "C:\Users\$env:USERNAME\OneDrive - Noble House Hotels & Resorts\Documents\Monthly Audits\2025 January Audits\January $propertyCode - O365 Users.xlsx"
    $sourceWorkbookPath = "C:\Users\$env:USERNAME\OneDrive - Noble House Hotels & Resorts\Documents\Monthly Audits\2025 January Audits\2025 January - O365 Users.xlsx"
    #$sourceWorkbookPath = "https://noblehousehotels44.sharepoint.com/sites/NHHROffice365Management/Shared Documents/Monthly Licensing Reports\O365 Users and Licenses.xlsx"

    try {
        Write-Host "Processing property code: $propertyCode"

        # Ensure the target directory exists
        $targetDirectory = [System.IO.Path]::GetDirectoryName($newWorkbookPath)
        if (-not (Test-Path -Path $targetDirectory)) {
            Write-Host "Creating directory: $targetDirectory"
            New-Item -ItemType Directory -Path $targetDirectory -Force
        }

        # Create a new workbook
        Write-Host "Creating new workbook at: $newWorkbookPath"
        Export-Excel -Path $newWorkbookPath -WorksheetName "Sheet1" -AutoSize -AutoFilter

        # Open the source workbook
        Write-Host "Opening source workbook: $sourceWorkbookPath"
        $excelSource = Open-ExcelPackage -Path $sourceWorkbookPath

        if ($excelSource -eq $null) {
            Write-Host "Failed to open source workbook: $sourceWorkbookPath"
            continue
        }

        $worksheetSource = $excelSource.Workbook.Worksheets[$propertyCode]

        if ($worksheetSource -eq $null) {
            Write-Host "Worksheet '$propertyCode' does not exist in the source workbook."
            Close-ExcelPackage -ExcelPackage $excelSource
            continue
        }

        # Open the destination workbook
        Write-Host "Opening destination workbook: $newWorkbookPath"
        $excelDestination = Open-ExcelPackage -Path $newWorkbookPath

        if ($excelDestination -eq $null) {
            Write-Host "Failed to open destination workbook: $newWorkbookPath"
            Close-ExcelPackage -ExcelPackage $excelSource
            continue
        }

        # Remove the existing property sheet if it exists
        $removeSheet = $excelDestination.Workbook.Worksheets[$propertyCode]
        if ($removeSheet -ne $null) {
            Write-Host "Removing existing worksheet: $propertyCode"
            $excelDestination.Workbook.Worksheets.Delete($removeSheet)
        }

        # Copy the worksheet from the source to the destination workbook
        Write-Host "Copying worksheet: $propertyCode"
        $excelDestination.Workbook.Worksheets.Add($propertyCode, $worksheetSource)

        # Remove "Sheet1" from the destination workbook
        $sheetToDelete = $excelDestination.Workbook.Worksheets["Sheet1"]
        if ($sheetToDelete -ne $null) {
            Write-Host "Removing worksheet: Sheet1"
            $excelDestination.Workbook.Worksheets.Delete($sheetToDelete)
        }

        # Save and close the destination workbook
        Write-Host "Saving destination workbook"
        Close-ExcelPackage -ExcelPackage $excelDestination
    }
    catch {
        Write-Host "Error processing $($propertyCode): $($_.Exception.Message)"
    }
    finally {
        # Ensure the Excel packages are closed
        if ($null -ne $excelSource) {
            Write-Host "Closing source workbook"
            Close-ExcelPackage -ExcelPackage $excelSource
        }
        if ($null -ne $excelDestination) {
            Write-Host "Closing destination workbook"
            Close-ExcelPackage -ExcelPackage $excelDestination
        }
    }
}
