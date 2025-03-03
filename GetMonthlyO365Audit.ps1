function Get-CurrentYearMonth {
    $currentDate = Get-Date

    $monthName = $currentDate.ToString("MMMM")
    $year = $currentDate.Year

    return [PSCustomObject]@{
        Month = $monthName
        Year  = $year
    }
}

function Subtract-Month {
    param (
        [int]$MonthsToSubtract = 1
    )

    # Get the current date
    $currentDate = Get-Date

    # Subtract the specified number of months
    $newDate = $currentDate.AddMonths(-$MonthsToSubtract)

    # Extract the new month name and year
    $monthName = $newDate.ToString("MMMM")
    $year = $newDate.Year

    # Return the new month name and year as a custom object
    return [PSCustomObject]@{
        Month = $monthName
        Year  = $year
    }
}

# ----------------------------------------------------------------------------------------------------------------------------------------------------------- #
# --------------------------------------------------- PREDEFINED VALUES ------------------------------------------------------------------------------------- #


# Dictionary mapping SkuId to license names
$skuIdMap = @{
    "4b585984-651b-448a-9e53-3b10f069cf7f" = "F3"
	"18181a46-0d4e-45cd-891e-60aabd171b4e" = "E1"
	"6fd2c87f-b296-42f0-b197-1e91e994b900" = "E3"
	"cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "BP"
    "e1b7c739-53a0-4d18-8c17-4b176a8e76f0" = "E5"
	"f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "BS"
    # Add more SkuId to license name mappings as needed
}

$additionalMap = @{
	"f30db892-07e9-47e9-837c-80727f46fd3d" = "Power Automate Free"
	"1c27243e-fb4d-42b1-ae8c-fe25c9616588" = "MS Teams Audio (Free)"
	"2b317a4a-77a6-4188-9437-b68a77b4e2c6" = "MS Intune Plan 1"
	"3f9f06f5-3c31-472c-985f-62d9c10ec167" = "Power Pages vTrial"
}

$propertyMapping = @{
	"ARG" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Argonaut"
        "Company Name" = "Argonaut Hotel"
        "Office Code" = "ARG"
        "Street" = "495 Jefferson St"
        "City" = "San Francisco"
        "State" = "CA"
    	"Zip" = "94109"
    	"Country" = "United States"
    	"Main Phone" = "415-563-0800"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("ARG")
    	"Audit Email" = "argo365audit@noblehousehotels.com"
	}
	"CAB" = @{
		    	"Addresses" = @(
        	@{
		"Short Name" = "Cabo"
    	"Company Name" = "Corazon Cabo"
    	"Office Code" = "CAB"
    	"Street" = "Av Pescadores S/N, Medano Beach"
    	"City" = "Cabo San Lucas"
    	"State" = "Baja California Sur"
    	"Zip" = "23453"
    	"Country" = "Mexico"
    	"Main Phone" = ""
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("CAB")
    	"Audit Email" = "cabo365audit@noblehousehotels.com"
	}
	"CHA" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Chatham"
    	"Company Name" = "Chatham Inn"
    	"Office Code" = "CHA"
    	"Street" = "359 Main Street"
    	"City" = "Chatham"
   		"State" = "MA"
    	"Zip" = "02633"
    	"Country" = "United States"
    	"Main Phone" = "508-945-9232"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("CHA")
    	"Audit Email" = "chio365audit@noblehousehotels.com"
	}
	"EDG" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Edgewater"
    	"Company Name" = "The Edgewater Hotel"
    	"Office Code" = "EDG"
    	"Street" = "2411 Alaskan Way"
    	"City" = "Seattle"
    	"State" = "WA"
    	"Zip" = "98120"
    	"Country" = "United States"
    	"Main Phone" = "206-792-5959"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("EDG")
    	"Audit Email" = "edgo365audit@noblehousehotels.com"
	}
	"ELJ" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Estancia"
    	"Company Name" = "Estancia La Jolla"
    	"Office Code" = "ELJ"
    	"Street" = "9700 N Torrey Pines Rd"
    	"City" = "La Jolla"
    	"State" = "CA"
    	"Zip" = "92037"
    	"Country" = "United States"
    	"Main Phone" = "855-430-7503"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("ELJ")
    	"Audit Email" = "eljo365audit@noblehousehotels.com"
	}
	"GWC" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Gateway Canyons"
    	"Company Name" = "Gateway Canyons Resort & Spa"
    	"Office Code" = "GWC"
    	"Street" = "43200 CO-141"
    	"City" = "Gateway"
    	"State" = "CO"
    	"Zip" = "81522"
    	"Country" = "United States"
    	"Main Phone" = "970-730-2058"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("GWC")
    	"Audit Email" = "gwco365audit@noblehousehotels.com"
	}
	"IOF" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Inn on Fifth"
    	"Company Name" = "Inn on Fifth"
    	"Office Code" = "IOF"
    	"Street" = "699 5th Ave S."
    	"City" = "Naples"
    	"State" = "FL"
    	"Zip" = "34102"
    	"Country" = "United States"
    	"Main Phone" = "239-403-8777"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("IOF")
    	"Audit Email" = "iofo365audit@noblehousehotels.com"
	}
	"JICR" = @{
		    	"Addresses" = @(
        	@{
	    "Short Name" = "Jekyll Club"
    	"Company Name" = "Jekyll Island Club Resort"
    	"Office Code" = "JICR"
    	"Street" = "371 N Riverview Dr."
    	"City" = "Jekyll Island"
    	"State" = "GA"
    	"Zip" = "31527"
    	"Country" = "United States"
    	"Main Phone" = "912-635-2600"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("JICR")
    	"Audit Email" = "jicro365audit@noblehousehotels.com"
	}
	"KKR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Kona Kai"
    	"Company Name" = "Kona Kai Resort & Spa"
    	"Office Code" = "KKR"
   		"Street" = "1551 Shelter Island Dr"
    	"City" = "San Diego"
    	"State" = "CA"
    	"Zip" = "92106"
    	"Country" = "United States"
    	"Main Phone" = "619-452-3138"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("KKR", "KKM")
    	"Audit Email" = "kkro365audit@noblehousehotels.com"
	}
	"LAP" = @{
    	"Addresses" = @(
        	@{
            	"Short Name"   = "Laplaya"
            	"Company Name" = "Laplaya Beach & Golf Resort"
            	"Office Code"  = "LAP"
            	"Street"       = "9891 Gulf Shore Dr"
            	"City"         = "Naples"
            	"State"        = "FL"
            	"Zip"          = "34108"
            	"Country"      = "United States"
            	"Main Phone"   = "239-597-3123"
            	"800 Number"   = ""
        	},
        	@{
            	"Short Name"   = "Laplaya Club"
            	"Company Name" = "Laplaya Beach & Golf Club"
            	"Office Code"  = "LAP"
            	"Street"       = "9891 Gulf Shore Dr"
            	"City"         = "Naples"
            	"State"        = "FL"
            	"Zip"          = "34108"
            	"Country"      = "United States"
            	"Main Phone"   = "239-254-5008"
            	"800 Number"   = ""
        	}
    	)
    	"Search Office Codes" = @("LAP")
    	"Audit Email" = "lapo365audit@noblehousehotels.com"
	}
	"LDM" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "L'Auberge"
    	"Company Name" = "L'Auberge Del Mar"
    	"Office Code" = "LDM"
    	"Street" = "1540 Camino Del Mar"
   		"City" = "Del Mar"
    	"State" = "CA"
    	"Zip" = "92014"
    	"Country" = "United States"
    	"Main Phone" = "858-386-1336"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("LDM")
    	"Audit Email" = "ldmo365audit@noblehousehotels.com"
	}
	"LPI" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Little Palm Island"
    	"Company Name" = "Little Palm Island Resort & Spa"
    	"Office Code" = "LPI"
    	"Street" = "28500 Overseas Hwy"
    	"City" = "Little Torch Key"
    	"State" = "FL"
    	"Zip" = "33042"
    	"Country" = "United States"
   		"Main Phone" = "305-684-8341"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("LPI")
    	"Audit Email" = "lpio365audit@noblehousehotels.com"
	}
	"MAR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Marquesa"
    	"Company Name" = "Marquesa Hotel"
    	"Office Code" = "MAR"
    	"Street" = "600 Fleming St."
    	"City" = "Key West"
    	"State" = "FL"
    	"Zip" = "33040"
    	"Country" = "United States"
    	"Main Phone" = "305-292-1919"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("MAR")
    	"Audit Email" = "maro365audit@noblehousehotels.com"
	}
	"MBR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Mission Bay Resort"
    	"Company Name" = "Mission Bay Resort"
    	"Office Code" = "MBR"
    	"Street" = "1775 E Mission Bay Dr"
    	"City" = "San Diego"
    	"State" = "CA"
    	"Zip" = "92109"
    	"Country" = "United States"
    	"Main Phone" = "619-276-4010"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("MBR")
    	"Audit Email" = "mbro365audit@noblehousehotels.com"
	}
	"NHHR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Noble House"
    	"Company Name" = "Noble House Hotels & Resorts"
    	"Office Code" = "NHHR"
    	"Street" = "600 6th ST S"
    	"City" = "Kirkland"
    	"State" = "WA"
    	"Zip" = "98033"
    	"Country" = "United States"
    	"Main Phone" = "425-827-8737"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("NHHR", "VENDOR")
    	"Audit Email" = " "
	}
	"NVWT" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Napa Valley Wine Train"
    	"Company Name" = "Napa Valley Wine Train"
    	"Office Code" = "NVWT"
    	"Street" = "1275 McKinstry"
    	"City" = "Napa"
    	"State" = "CA"
    	"Zip" = "94559"
    	"Country" = "United States"
    	"Main Phone" = "707-253-2111"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("NVWT")
    	"Audit Email" = "nvwto365audit@noblehousehotels.com"
	}
	"OKR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Ocean Key"
    	"Company Name" = "Ocean Key Resort & Spa"
    	"Office Code" = "OKR"
    	"Street" = "0 Duval Street"
    	"City" = "Key West"
    	"State" = "FL"
    	"Zip" = "33040"
    	"Country" = "United States"
    	"Main Phone" = "305-809-8072"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("OKR")
    	"Audit Email" = "okro365audit@noblehousehotels.com"
	}
	"HCA" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Harts Camp"
    	"Company Name" = "Harts Camp RV Park"
    	"Office Code" = "HCA"
    	"Street" = "33145 Webb Park Rd."
    	"City" = "Pacific City"
    	"State" = "OR"
    	"Zip" = "97135"
    	"Country" = "United States"
    	"Main Phone" = "503-965-7006"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("HCA")
    	"Audit Email" = "hcao365audit@noblehousehotels.com"
	}
	"CKA" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Inn at Cape Kiwanda"
    	"Company Name" = "Inn at Cape Kiwanda"
    	"Office Code" = "CKA"
    	"Street" = "33105 Cape Kiwanda Drive"
    	"City" = "Pacific City"
    	"State" = "OR"
    	"Zip" = "97135"
    	"Country" = "United States"
    	"Main Phone" = "888-965-7001"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("CKA")
    	"Audit Email" = "ckao365audit@noblehousehotels.com"
	}
	"HCL" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Headlands"
    	"Company Name" = "Headlands Costal Lodge"
    	"Office Code" = "HCL"
    	"Street" = "33000 Cape Kiwanda Drive"
    	"City" = "Pacific City"
    	"State" = "OR"
    	"Zip" = "97135"
    	"Country" = "United States"
    	"Main Phone" = "503-483-3000"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("HCL", "PCL", "PCH")
    	"Audit Email" = "hclo365audit@noblehousehotels.com"
	}
	"PBR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Pelican"
    	"Company Name" = "Pelican Grand Beach Resort"
    	"Office Code" = "PBR"
    	"Street" = "2000 N Ocean Blvd"
    	"City" = "Fort Lauderdale"
    	"State" = "FL"
    	"Zip" = "33305"
    	"Country" = "United States"
    	"Main Phone" = "(850) 654-1425"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("PBR")
    	"Audit Email" = "pbro365audit@noblehousehotels.com"
	}
	"POK" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Pokeo"
    	"Company Name" = "PokeoKeo"
    	"Office Code" = "POK"
    	"Street" = "600 6th ST S"
    	"City" = "Kirkland"
    	"State" = "WA"
    	"Zip" = "98033"
    	"Country" = "United States"
    	"Main Phone" = "877-662-5387"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("POK")
    	"Audit Email" = " "
	}
	"POR" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Portofino"
    	"Company Name" = "The Portofino Hotel & Marina"
    	"Office Code" = "POR"
    	"Street" = "260 Portofino Way"
    	"City" = "Redondo Beach"
    	"State" = "CA"
    	"Zip" = "90277"
    	"Country" = "United States"
    	"Main Phone" = "310-421-4195"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("POR")
    	"Audit Email" = "poro365audit@noblehousehotels.com"
	}
	"RTI" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "River Terrace Inn"
    	"Company Name" = "River Terrace Inn"
    	"Office Code" = "RTI"
    	"Street" = "1600 Soscol Ave"
    	"City" = "Napa"
    	"State" = "CA"
    	"Zip" = "94559"
    	"Country" = "United States"
    	"Main Phone" = "707-927-2217"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("RTI")
    	"Audit Email" = "rtio365audit@noblehousehotels.com"
	}
	"SOL" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Sole"
    	"Company Name" = "Sole Miami"
    	"Office Code" = "SOL"
    	"Street" = "17315 Collins Ave"
    	"City" = "Sunny Isles Beach"
    	"State" = "FL"
    	"Zip" = "33160"
    	"Country" = "United States"
    	"Main Phone" = "786-374-2211"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("SOL")
    	"Audit Email" = "solo365audit@noblehousehotels.com"
	}
	"SRC" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Snake River"
    	"Company Name" = "Snake River Sporting Club"
    	"Office Code" = "SRC"
    	"Street" = "14885 Sporting Club Road"
    	"City" = "Jackson"
    	"State" = "WY"
    	"Zip" = "83001"
    	"Country" = "United States"
    	"Main Phone" = "(307) 733-3444"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("SRC")
    	"Audit Email" = "srco365audit@noblehousehotels.com"
	}
	"TR" = @{
    	"Addresses" = @(
        	@{
            	"Short Name"   = "Teton Mountain Lodge"
            	"Company Name" = "Teton Mountain Lodge & Spa"
            	"Office Code"  = "TR - TML"
            	"Street"       = "3385 Cody Ln"
            	"City"         = "Teton Village"
            	"State"        = "WY"
            	"Zip"          = "83025"
            	"Country"      = "United States"
            	"Main Phone"   = "307-201-6066"
            	"800 Number"   = ""
        	},
        	@{
            	"Short Name"   = "Hotel Terra"
            	"Company Name" = "Hotel Terra Jackson Hole"
            	"Office Code"  = "TR - HTJ"
            	"Street"       = "3335 Village Dr"
            	"City"         = "Teton Village"
            	"State"        = "WY"
            	"Zip"          = "83025"
            	"Country"      = "United States"
            	"Main Phone"   = "307-201-6065"
            	"800 Number"   = ""
        	}
    	)
    	"Search Office Codes" = @("TR")
    	"Audit Email" = "tro365audit@noblehousehotels.com"
}
	"TSH" = @{
	    	"Addresses" = @(
        	@{
	    "Short Name" = "Stella"
    	"Company Name" = "The Stella Hotel"
    	"Office Code" = "TSH"
    	"Street" = "4100 Lake Atlas Dr"
    	"City" = "Bryan"
    	"State" = "TX"
   		"Zip" = "77807"
   		"Country" = "United States"
    	"Main Phone" = "979-421-4025"
    	"800 Number" = ""
    	        	}
    	)
    	"Search Office Codes" = @("TSH")
    	"Audit Email" = "tsho365audit@noblehousehotels.com"
	}
}

$allUsers = @()
$dateInfo = Get-CurrentYearMonth
$currentDate = Get-Date -Format "dd.MM.HH.mm"
$folderPath = "C:\Users\$env:USERNAME\OneDrive - Noble House Hotels & Resorts\Documents\Monthly Audits\$($dateInfo.Year) $($dateInfo.Month) Audits"
$filePath = Join-Path -Path $folderPath -ChildPath "$($dateInfo.Year) $($dateInfo.Month) - O365 Users.xlsx"
$newFileName = "$($dateInfo.Year) $($dateInfo.Month) - O365 Users - OLD ($currentDate).xlsx"
$newFilePath = Join-Path -Path $folderPath -ChildPath $newFileName

# Check if the folder exists
if (!(Test-Path -Path $folderPath)) {
    # Folder does not exist, create it
    New-Item -ItemType Directory -Path $folderPath | Out-Null
    Write-Host -ForegroundColor Green "Folder created: $folderPath"
} else {
    # Folder exists, check if the file exists within it
    if (Test-Path -Path $filePath) {
        # File exists, rename it
        Rename-Item -Path $filePath -NewName $newFilePath
        Write-Host -ForegroundColor Green "File renamed to: $newFilePath"
    } else {
        # File does not exist, do nothing
        Write-Host -ForegroundColor Yellow "Folder exists, but file not found: $filePath"
    }
}
$workbookPath = "C:\Users\$env:USERNAME\OneDrive - Noble House Hotels & Resorts\Documents\Monthly Audits\$($dateInfo.Year) $($dateInfo.Month) Audits\$($dateInfo.Year) $($dateInfo.Month) - O365 Users.xlsx"


# ----------------------------------------------------------- END PREDEFINED VALUES -------------------------------------------------------------------------------- #

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
# Retrieve users with specified properties and filters
$users = Get-MgUser -All -Property CompanyName, AssignedLicenses, OfficeLocation, DisplayName, UserPrincipalName, JobTitle, Department, AccountEnabled
$users | Where-Object {
		($_.AssignedLicenses | Where-Object { $skuIdMap.ContainsKey($_.SkuId) }).Count -gt 0
    } | ForEach-Object {
		$licenses = @()
		$additionallicenses = @()
		foreach ($license in $_.AssignedLicenses) {
			if ($skuIdMap.ContainsKey($license.SkuId)) {
				$licenses += $skuIdMap[$license.SkuId]
			} else {
				if ($additionalMap.ContainsKey($license.SkuId)) {
					$additionallicenses += $additionalMap[$license.SkuId]
				} else {
					$additionallicenses += $license.SkuId
				}
			}
		}
		[PSCustomObject]@{
			Office    			= $_.OfficeLocation
			Licenses  			= $licenses -join "; "
			DisplayName      = $_.DisplayName
			UserPrincipalName 	= $_.UserPrincipalName
			Title          		= $_.JobTitle
			Comments 			= ""
		}
	} | Export-Excel -Path $workbookPath -WorksheetName "All Licensed Users" -TableName "All_Users" -AutoSize -AutoFilter


foreach ($propertyCode in $propertyMapping.Keys) {

	$filteredUsers = @()
	$property = $propertyMapping[$propertyCode]
	$searchOfficeCodes = $property["Search Office Codes"]

    foreach ($searchCode in $searchOfficeCodes) {
        $filteredUsers += $users | Where-Object {
            ($_.OfficeLocation -like "$searchCode*") -and
            ($_.AssignedLicenses | Where-Object { $skuIdMap.ContainsKey($_.SkuId) }).Count -gt 0
        }
    }

	# Process each user to extract license details and convert SkuId to license names
	$processedUsers = $filteredUsers | ForEach-Object {
		$licenses = @()
		$additionallicenses = @()
		foreach ($license in $_.AssignedLicenses) {
			if ($skuIdMap.ContainsKey($license.SkuId)) {
				$licenses += $skuIdMap[$license.SkuId]
			} else {
				if ($additionalMap.ContainsKey($license.SkuId)) {
					$additionallicenses += $additionalMap[$license.SkuId]
				} else {
					$additionallicenses += $license.SkuId
				}
			}
		}
		[PSCustomObject]@{
			Office    			= $_.OfficeLocation
			Licenses  			= $licenses -join "; "
			DisplayName      = $_.DisplayName
			UserPrincipalName 	= $_.UserPrincipalName
			Title          		= $_.JobTitle
			Comments 			= ""
		}
	}

    # Export the processed users to the specified location in the workbook
	# $processedUsers | Export-Excel -Path "C:\Users\JakeRobinson\OneDrive - Noble House Hotels & Resorts\Desktop\O365 Users and Licenses New2.xlsx" -WorksheetName "$propertyCode" -AutoSize -AutoFilter -TableName "$propertyCode" -StartRow 10
	# Load the existing workbook
$excelPackage = Open-ExcelPackage -Path $workbookPath
$worksheet = $excelPackage.Workbook.Worksheets.Add($propertyCode)

# Define the headers
$headers = @("Short Name", "Company Name", "Office Code", "Street", "City", "State", "Zip", "Country", "Main Phone", "800 Number")

	Set-ExcelRange -Worksheet $worksheet -Range "A1" -Value $headers[0] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "B1" -Value $headers[1] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "C1" -Value $headers[2] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "D1" -Value $headers[3] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "E1" -Value $headers[4] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "F1" -Value $headers[5] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "G1" -Value $headers[6] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "H1" -Value $headers[7] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "I1" -Value $headers[8] -Bold
	Set-ExcelRange -Worksheet $worksheet -Range "J1" -Value $headers[9] -Bold




$row = 2
$property = $propertyMapping[$propertyCode]
$addresses = $property["Addresses"]
	Set-ExcelRange -Worksheet $worksheet -Range "I6" -Value $property["Audit Email"]
foreach ($address in $addresses) {
    Set-ExcelRange -Worksheet $worksheet -Range "A$row" -Value $address["Short Name"]
    Set-ExcelRange -Worksheet $worksheet -Range "B$row" -Value $address["Company Name"]
    Set-ExcelRange -Worksheet $worksheet -Range "C$row" -Value $address["Office Code"]
    Set-ExcelRange -Worksheet $worksheet -Range "D$row" -Value $address["Street"]
    Set-ExcelRange -Worksheet $worksheet -Range "E$row" -Value $address["City"]
    Set-ExcelRange -Worksheet $worksheet -Range "F$row" -Value $address["State"]
    Set-ExcelRange -Worksheet $worksheet -Range "G$row" -Value $address["Zip"]
    Set-ExcelRange -Worksheet $worksheet -Range "H$row" -Value $address["Country"]
    Set-ExcelRange -Worksheet $worksheet -Range "I$row" -Value $address["Main Phone"]
    Set-ExcelRange -Worksheet $worksheet -Range "J$row" -Value $address["800 Number"]
    # Move to the next row for the next address or property
    $row++
}

$row++
Set-ExcelRange -Worksheet $worksheet -Range "B$row" -Value "BS" -Bold
Set-ExcelRange -Worksheet $worksheet -Range "C$row" -Value "BP" -Bold
Set-ExcelRange -Worksheet $worksheet -Range "D$row" -Value "F3" -Bold
Set-ExcelRange -Worksheet $worksheet -Range "E$row" -Value "E1" -Bold
Set-ExcelRange -Worksheet $worksheet -Range "F$row" -Value "E3" -Bold
Set-ExcelRange -Worksheet $worksheet -Range "G$row" -Value "E5" -Bold

# Insert and format ANNUAL row
$row++
Set-ExcelRange -Worksheet $worksheet -Range "A$row" -Value "ANNUAL" -Bold
$worksheet.Cells.Item($row, 2).Formula = "=COUNTIF(B11:B5000,""BS ANNUAL"")"
$worksheet.Cells.Item($row, 3).Formula = "=COUNTIF(B11:B5000,""BP ANNUAL"")"
$worksheet.Cells.Item($row, 4).Formula = "=COUNTIF(B11:B5000,""F3 ANNUAL"")"
$worksheet.Cells.Item($row, 5).Formula = "=COUNTIF(B11:B5000,""E1 ANNUAL"")"
$worksheet.Cells.Item($row, 6).Formula = "=COUNTIF(B11:B5000,""E3 ANNUAL"")"
$worksheet.Cells.Item($row, 7).Formula = "=COUNTIF(B11:B5000,""E5 ANNUAL"")"


# Insert and format Month to Month row
$row++
Set-ExcelRange -Worksheet $worksheet -Range "A$row" -Value "Month to Month" -Bold
$worksheet.Cells.Item("B$row").Formula = "=COUNTIF(B11:B5000,""BS"")"
$worksheet.Cells.Item("C$row").Formula = "=COUNTIF(B11:B5000,""BP"")"
$worksheet.Cells.Item("D$row").Formula = "=COUNTIF(B11:B5000,""F3"")"
$worksheet.Cells.Item("E$row").Formula = "=COUNTIF(B11:B5000,""E1"")"
$worksheet.Cells.Item("F$row").Formula = "=COUNTIF(B11:B5000,""E3"")"
$worksheet.Cells.Item("G$row").Formula = "=COUNTIF(B11:B5000,""E5"")"

# Insert and format Total row
$row++
$firstRow = $row - 1
$secondRow = $row - 2
Set-ExcelRange -Worksheet $worksheet -Range "A$row" -Value "Total" -Bold
$worksheet.Cells.Item("B$row").Formula = "=SUM(" + "B" + $secondRow + ":" + "B" + $firstRow + ")"
$worksheet.Cells.Item("C$row").Formula = "=SUM(" + "C" + $secondRow + ":" + "C" + $firstRow + ")"
$worksheet.Cells.Item("D$row").Formula = "=SUM(" + "D" + $secondRow + ":" + "D" + $firstRow + ")"
$worksheet.Cells.Item("E$row").Formula = "=SUM(" + "E" + $secondRow + ":" + "E" + $firstRow + ")"
$worksheet.Cells.Item("F$row").Formula = "=SUM(" + "F" + $secondRow + ":" + "F" + $firstRow + ")"
$worksheet.Cells.Item("G$row").Formula = "=SUM(" + "G" + $secondRow + ":" + "G" + $firstRow + ")"

$row++


	# Define the starting cell for the new table
	$startRow = "10"

		# Export the processed users to the specified location in the workbook
		try {
			$processedUsers | Export-Excel -ExcelPackage $excelPackage -WorksheetName $propertyCode -StartRow $startRow -TableName $propertyCode -AutoSize -AutoFilter -FreezePane 0, 10 -PassThru | Out-Null
		} catch {
			if (!$processedUsers) {
				Write-Host -f Red "No User's found at $propertyCode"
			} else {
				Write-Error "Failed to export users to worksheet ""$propertyCode"": $_"
			}
		}
		$allUsers += $processedUsers
		# Attempt to auto-size columns
		try {
			$worksheet.Cells.AutoFitColumns()
		} catch {
			Write-Error "Failed to auto-size columns: $_"
		}


	# Save and close the workbook
	try {
		Close-ExcelPackage $excelPackage
		$workbook = $excel.Workbooks.Open($workbookPath)
		$worksheet = $workbook.Worksheets.Item($propertyCode)
		$worksheet.Activate()
		$worksheet.Application.ActiveWindow.SplitRow = 10
		$worksheet.Application.ActiveWindow.FreezePanes = $true
		$workbook.Save()
		$workbook.Close()
	} catch {
		Write-Error "Error saving file: $_"
	}
}
	$workbook = $excel.Workbooks.Open($workbookPath)
    $worksheets = $workbook.Sheets

    try {
        # Sort worksheets alphabetically
        $worksheetNames = @()
        foreach ($sheet in $worksheets) {
            $worksheetNames += $sheet.Name
        }

        $worksheetNames = $worksheetNames | Sort-Object

        for ($i = 0; $i -lt $worksheetNames.Count; $i++) {
            $sheet = $worksheets.Item($worksheetNames[$i])
            $sheet.Move([System.Reflection.Missing]::Value, $workbook.Sheets.Item($i + 1))
        }

        # Save and close the workbook
        $workbook.Save()
    } finally {
        $workbook.Close()
        $excel.Quit()
        }

<#
        $newWorkbookPath = "C:\Users\$env:USERNAME\Documents\O365Users-Octo.xlsx"
    	$sourceWorkbookPath = ""
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
 #>
