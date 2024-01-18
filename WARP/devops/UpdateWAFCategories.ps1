#requires -Version 7
<#
.SYNOPSIS
    Uses the published Well Architected Framework documentation to generate category names and descriptions for the WAF Category file. 
    The script preserves existing descriptions by default. To update existing descriptions run with the -Force parameter or delete the existing descriptions.
    The script requires an internet connection and uses content from:
    https://github.com/MicrosoftDocs/well-architected

.PARAMETER Force
    The "Force" parameter will overwrite existing descriptions in the csv file.

.INPUTS
    WAF Category Descriptions.csv file which must be in the same directory as this script.
    If the file is not present, it will be created.

.OUTPUTS
    Updated WAF Category Descriptions.csv file

.NOTES
    You must be connected to the internet to run this script. 
    It takes a dependency on the Well-Architected documentation repo.

.LINK
    https://github.com/Azure/WellArchitected-Tools/

#>

[CmdletBinding()]
param (
    # Indicates CSV file for input
    [Parameter()][string]
    $WellArchitectedDocsRepo = "https://raw.githubusercontent.com/MicrosoftDocs/well-architected/main/well-architected/",

    [Parameter()][string]
    $WAFCategoryFileName = "WAF Category Descriptions.csv",

    [Parameter()]
    [switch] $Force
)

# Get the working directory from the script
$workingDirectory = (Get-Location).Path 

Function EnsureRepoFile($fileName, $remoteFile) {
    if([System.IO.File]::Exists($workingDirectory + "\" + $fileName)) {
        #check if file has been updated in the last day
        $fileAge = (Get-Date) - (Get-Item -Path ($workingDirectory + "\" + $fileName)).LastWriteTime
        if ($fileAge.Days -gt 1) {
            #file is older than a day, so update it
            Remove-Item -Path ($workingDirectory + "\" + $fileName) -Force
        }

    }

    if(![System.IO.File]::Exists($workingDirectory + "\" + $fileName)) {
        $toc = Invoke-RestMethod -Uri ($WellArchitectedDocsRepo + $remoteFile)
        $toc | Out-File -FilePath ($workingDirectory + "\" + $fileName)
    }
}

Function GetWAF2CategoryCaption($lineData) {
    $searchCode = $lineData.Substring(0,5).Trim()
    $toc = ""
    $categoryCaption = $searchCode

    EnsureRepoFile -fileName "WAF-toc.yml" -remoteFile "TOC.yml"

    $toc = Get-Content -Path ($workingDirectory + "\WAF-toc.yml")

    foreach($line in $toc) {
        if($line.Trim().StartsWith("- name: " + $searchCode)) {
            if($categoryCaption.Length -gt 5) {
                $categoryCaption = $categoryCaption + ","
            }
            $categoryArray = $line.Trim().Split(":")
            $categoryCaption = $categoryCaption + $categoryArray[$categoryArray.Length - 1].Substring(2)
        }
    }

    return $categoryCaption

}

Function GetWAF2CategoryDescription($lineData) {
    $searchCode = $lineData.Substring(0,5).Trim()
    $toc = ""
    $description = $null

    EnsureRepoFile -fileName "WAF-toc.yml" -remoteFile "TOC.yml"

    $toc = Get-Content -Path ($workingDirectory + "\WAF-toc.yml")

    foreach($line in $toc) {
        if($line.Contains("checklist.md")) {
            $array = $line.Split(":");
            $filename = $array[$array.Length - 1].Trim();
            $remoteFileName = $filename;
            $localFileName = $filename.Replace("/", "-");

            EnsureRepoFile -fileName $localFileName -remoteFile $remoteFileName

            $checklist = Get-Content -Path ($workingDirectory + "\" + $localFileName)

            foreach($checklistLine in $checklist) {
                if($checklistLine.Contains($searchCode)) {
                    $checklistArray = $checklistLine.Split("|")
                    $descriptionRaw = $checklistArray[$checklistArray.Length - 2].Trim()
                    $description = $descriptionRaw.Replace("**", "")
                    break
                }   
            }
        }

        if($description -ne $null) {
            break
        }
    }

    return $description
}

# Check that WAF Descriptions File exists
if(![System.IO.File]::Exists($workingDirectory + "\" + $WAFCategoryFileName)) {
    # Create the file
    $header = "Pillar,Category,Caption,Description"
    $header | Out-File -FilePath ($workingDirectory + "\" + $WAFCategoryFileName)
}

# Get existing content from CSV file
$existingContent = @(Import-Csv -Path ($workingDirectory + "\" + $WAFCategoryFileName))

EnsureRepoFile -fileName "WAF-toc.yml" -remoteFile "TOC.yml"

# Go through the TOC and find category indicators
$toc = Get-Content -Path ($workingDirectory + "\WAF-toc.yml")

foreach($line in $toc) {
    $pillar = "";
    $category = "";
    $caption = $null;
    $description = $null;

    if($line.Trim().StartsWith("- name: ")) {
        $lineArray = $line.Trim().Split(":")
        if($lineArray.Length -ne 3) {
            # not the type of line we're looking for
            continue
        }

        $pillarIndicator = $lineArray[1].Trim();
        $recommendationIndicator = $lineArray[2].Trim().Substring(0,2);
        if($pillarIndicator.Length -ne 2) {
            # not the type of line we're looking for
            continue
        }

        if($pillarIndicator.Equals("CO")) {
            $pillar = "Cost Optimization"
        } elseif ($pillarIndicator.Equals("OE")) {
            $pillar = "Operational Excellence"
        } elseif ($pillarIndicator.Equals("PE")) {
            $pillar = "Performance Efficiency"
        } elseif ($pillarIndicator.Equals("RE")) {
            $pillar = "Reliability"
        } elseif ($pillarIndicator.Equals("SE")) {
            $pillar = "Security"
        } else {
            # not the type of line we're looking for
            continue
        }

        $category = $pillarIndicator + ":" + $recommendationIndicator

        $caption = GetWAF2CategoryCaption -lineData $category
        $description = GetWAF2CategoryDescription -lineData $category

        if($caption -eq $null) {
            # this hasn't worked - let's skip
            continue
        }

        if($description -eq $null) {
            # this hasn't worked - let's skip
            continue
        }

        $exists = $false
        # does this category already exist
        $existingContent | Where-Object {$_.Category -eq $category} | ForEach-Object {
            if($Force) {
                # overwrite existing description
                $_.Pillar = $pillar
                $_.Caption = $caption
                $_.Description = $description
            }

            $exists = $true
        }

        if(!$exists) {
            # add new category
            $newLine = [PSCustomObject] @{
                'Pillar' = $null
                'Category' = $null
                'Caption' = $null
                'Description' = $null
            }
            $newLine.Pillar = $pillar
            $newLine.Category = $category
            $newLine.Caption = $caption
            $newLine.Description = $description

            $existingContent += $newLine
        }
    }
}

Remove-Item -Path ($workingDirectory + "\" + $WAFCategoryFileName) -Force
$existingContent | Export-Csv -Path ($workingDirectory + "\" + $WAFCategoryFileName) -NoTypeInformation -UseQuotes AsNeeded