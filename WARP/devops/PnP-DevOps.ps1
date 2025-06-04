<#
.SYNOPSIS
    Creates epics and issues in an Azure DevOps project based on Well-Architected and other Azure assessment findings .csv file.
    
.DESCRIPTION
    Creates epics and issues in an Azure DevOps project based on Well-Architected and other Azure assessment findings .csv file.

.PARAMETER DevOpsPersonalAccessToken
    Personal Access Token from Azure DevOps

.PARAMETER DevOpsProjectUri
    URI fo the Azure DevOps project
    
.PARAMETER AssessmentCsvPath
    .csv file from Well-Architected and other Azure assessment export

.PARAMETER DevOpsTagName
    Name of assessment Example: "WAF"

.PARAMETER DevOpsWorkItemType
    The type of DevOps work item to create and link to the Epics. Certain project types support certain work items. SCRUM(Feature), Agile(Feature & Issue), Basic(Issue)

.OUTPUTS
        Status message text

.EXAMPLE
    PnP-DevOps -DevOpsPersonalAccessToken xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -DevOpsProjectUri https://dev.azure.com/organization/project -DevOpsTagName WAF -DevOpsWorkItemType Feature -AssessmentCsvPath c:\temp\Azure_Well_Architected_Review_Jan_1_2023_1_00_00_PM.csv
    Adds the items from the Well-Architected and other Azure assessments .csv export to a DevOps project as work itmes.

.NOTES
    Make sure 'WAF Category Descriptions.csv' is in the same directory as this script. It is used for well-architected assessments to map the old category names to the new category names 

.LINK

#>

[CmdletBinding()]
param (
    [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DevOpsPersonalAccessToken,
    [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][uri]$DevOpsProjectUri,
    [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DevOpsTagName,
    [parameter(Mandatory = $true)][ValidateSet("Feature", "Issue")][string]$DevOpsWorkItemType,
    [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][System.IO.FileInfo]$AssessmentCsvPath
)

$ErrorActionPreference = "break"


#region Functions

# Get settings for either Azure DevOps
function Get-AdoSettings {

    $authHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($DevOpsPersonalAccessToken)")) }
    
    $uriBase = $DevOpsProjectUri.ToString().Trim("/") + "/"

    $settings = @{
        authHeader = $authHeader
        uriBase    = $uriBase
    }

    return $settings
}


$content = Get-Content  $AssessmentCsvPath

#Capturing first line of csv file to later check for "Well-Architected" string
$assessmentTypeCheck = ($content | Select-Object -First 1)

#Updated function to process import of items for non-Well-Architected assessments
function Import-AssessmentOther {
  
    $tableStartPattern = ($content | Select-String "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context" | Select-Object * -First 1)
    $tableStart = ( $tableStartPattern.LineNumber ) - 1
    $endStringIdentifier = $content | Where-Object { $_.Contains("--,,") } | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Hack to change blank reporting category to Azure Advisor
    $devOpsList | 
    Where-Object -Property ReportingCategory -eq "" | 
    ForEach-Object { $_.ReportingCategory = "Defender for Cloud" }


    # Get unique list of ReportCategory column
    $reportingCategories = @{}
    $devOpsList | 
    Select-Object -Property ReportingCategory -Unique | 
    Sort-Object  -Property ReportingCategory |
    ForEach-Object { 
        $reportingCategories[$_.ReportingCategory] = ""       
    }

    # Add Decription
    $devOpsList | Add-Member -Name Description -MemberType NoteProperty -Value ""

    $devOpsList | 
    ForEach-Object { 
        $_.Description = "<a href=`"$($_.Link)`">$($_.'Link-Text')</a>"
    }

    $assessment = @{
        name                = $DevOpsTagName
        reportingCategories = $reportingCategories
        recommendations     = $devOpsList
    }

    return $assessment
}
function Import-Assessment {

    [System.IO.FileInfo]$sourceScript = $PSCmdlet.MyInvocation.MyCommand.Source 
    $workingDirectory = $sourceScript.DirectoryName

    try {
        $descriptionsFile = Import-Csv -Path "$workingDirectory\WAF Category Descriptions.csv"
    }
    catch {
        Write-Error -Message "Unable to open $workingDirectory\WAF Category Descriptions.csv"
        exit
    }

   
    $tableStartPattern = ($content | Select-String "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context" | Select-Object * -First 1)
    $tableStart = ( $tableStartPattern.LineNumber ) - 1
    $endStringIdentifier = $content | Where-Object { $_.Contains("--,,") } | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Hack to change blank reporting category to Azure Advisor
    $devOpsList | 
    Where-Object -Property ReportingCategory -eq "" | 
    ForEach-Object { $_.ReportingCategory = "Defender for Cloud" }

    # Get unique list of ReportCategory column
   
    # Populate $categoryMapping and $reportingCategories
    $devOpsList | 
    Select-Object -Property ReportingCategory, Category -Unique | 
    Sort-Object -Property ReportingCategory |
    ForEach-Object {
        $currentReportingCategory = $_.ReportingCategory
        $currentPillar = $_.Category
        $categoryTitle = ($descriptionsFile | Where-Object { $_.Pillar -eq $currentPillar -and $_.Category.StartsWith($currentReportingCategory) }).Caption
        if (-not $categoryTitle) {
            $categoryTitle = $currentReportingCategory # Fallback to existing ReportingCategory if no mapping found
        }
        $categoryMapping[$currentReportingCategory] = $categoryTitle
        $reportingCategories[$currentReportingCategory] = $categoryTitle
    }


    # Add Decription
    $devOpsList | Add-Member -Name Description -MemberType NoteProperty -Value ""

    $devOpsList | 
    ForEach-Object { 
        $_.Description = "<a href=`"$($_.Link)`">$($_.'Link-Text')</a>"
    }

    # Setting up the $assessment object
    $assessment = @{
        name                = $DevOpsTagName
        reportingCategories = $reportingCategories
        recommendations     = $devOpsList
    }

    return $assessment
}


function GetMappedReportingCategory {
    <#
    .DESCRIPTION
    This function takes an old category name as input and returns a new category name. The category name is used as the epic name in Azure DevOps.
    It uses a mapping stored in a hashtable $categoryMapping to find and return the corresponding new category name.  $categoryMapping is build from WAF Category Description.csv
    If the old category name does not exist in the mapping, the function returns the old category name.
    #>    
    param (
        $reportingCategrory
    )

    $newReportingCategory = if ($null -ne $categoryMapping -and $categoryMapping.ContainsKey($reportingCategrory)) {  
        $categoryMapping[$reportingCategrory] # map the old category to the new category # map the old category to the new category
    }
    else {
        $reportingCategrory # no mapping found, keep the old category
    }

    return $newReportingCategory
}  
    


function Search-EpicsAdo {
    param (
        $settings
    )

    $foundEpics = @{}

    # Logic to check and update epics
    $body = "{
                    `"query`": `"SELECT [System.Id] FROM workitems WHERE [System.TeamProject] = @project AND [System.WorkItemType] = 'Epic' AND [System.State] <> ''`"}"

    try {
        $getQueryUri = $settings.uriBase + "_apis/wit/wiql?api-version=6.0-preview.2"
        $results = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $settings.authHeader -Body $body
        if ($results.workItems.Count -gt 0) {
            foreach ($epic in $results.workItems.id) {
                $getEpicQueryUri = $settings.uriBase + "_apis/wit/workitems/" + $epic + "?api-version=6.0-preview.2"
                $epicWorkItem = Invoke-RestMethod -Uri $getEpicQueryUri -Method GET -ContentType "application/json" -Headers $settings.authHeader
                $epicNameFromAdo = $epicWorkItem.fields.'System.Title'

                $foundEpics[$epicNameFromAdo] = $epicWorkItem.url
            }
        }
    }
    catch {
        Write-Error "Error while querying Azure DevOps for Epics: $($Error[0].Exception.ToString())"
        exit
    }


    return $foundEpics

}


#Create the Epic in DevOps per category/focus area
function Add-EpicAdo {
    param (
        $settings,
        $epicName
    )

    try {

        $body = "[
            {
                `"op`": `"add`",
                `"path`": `"/fields/System.Title`",
                `"value`": `"$($epicName)`"
            }
        ]"
        
        
        Write-Host "Adding new Epic to ADO: $epicName" 

        $postEpicUri = $settings.uriBase + "_apis/wit/workitems/$" + "Epic" + "?api-version=6.0-preview.2"
        $epicWorkItem = Invoke-RestMethod -Uri $postEpicUri -Method POST -ContentType "application/json-patch+json" -Headers $settings.authHeader -Body $body
        $epicUrl = $epicWorkItem.url

        return $epicUrl 

    }
    catch {
        Write-Output $epicUrl
        Write-Error "Error creating Epic in DevOps: $($Error[0].Exception.ToString())"
        exit
    }
}

#Retrieve all work items from DevOps
function Get-WorkItemsAdo {
    param (
        $settings
    )

    #Iterate entire devops issues
    $body = "{
    `"query`": `"Select * From WorkItems Where [Work Item Type] = '$DevOpsWorkItemType' AND [State] <> 'Closed' AND [State] <> 'Removed' AND [System.TeamProject] = @project order by [Microsoft.VSTS.Common.Priority] asc, [System.CreatedDate] desc`"
    }"

    $getQueryUri = $settings.uriBase + "_apis/wit/wiql?api-version=6.0-preview.2"
    $results = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $settings.authHeader -Body $body


    $workItemsAdo = @()
    try {
        #Gather details per devops item
        if ($results.workItems.Count -gt 0) {
            foreach ($wi in $results.workItems.id) {
                $getWIQueryUri = $settings.uriBase + "_apis/wit/workitems/" + $wi + "?api-version=6.0-preview.2"
                $workItem = Invoke-RestMethod -Uri $getWIQueryUri -Method GET -ContentType "application/json" -Headers $settings.authHeader
                $workItemsAdo += $workItem
            }
        }
        else {
            Write-Verbose "There are no work items of type Issue in DevOps yet"
        }
    }
    catch {
        Write-Error "Error while querying devops for work items: $($Error[0].Exception.ToString())"
    }

    return $workItemsAdo
}

#Insert Feature into DevOps
function Add-NewIssueToDevOps {
    param (
        $settings,
        $assessment,
        $Title,
        $Effort,
        $Tags,
        $Priority,
        $BusinessValue,
        $TimeCriticality,
        $Risk,
        $Description,
        $linkedItem
    )
   
    if ($Title -eq "" -or $null -eq $Title) { $Title = "NA" }
    if ($Effort -eq "" -or $null -eq $Effort) { $Effort = "0" }
    #if($Tags -eq "" -or $null -eq $Tags){$Tags="NA"}
    if ($Priority -eq "" -or $null -eq $Priority) { $Priority = "4" }
    if ($BusinessValue -eq "" -or $null -eq $BusinessValue) { $BusinessValue = "0" }
    if ($TimeCriticality -eq "" -or $null -eq $TimeCriticality) { $TimeCriticality = "0" }
    if ($Risk -eq "" -or $null -eq $Risk) { $Risk = "3 - Low" }
    if ($Description -eq "" -or $null -eq $Description) { $Description = "NA" }

    
    if ($Tags -eq "" -or $null -eq $Tags) {
        $Tags = $assessment.name
    }
    else {
        $Tags = @($Tags, $assessment.name) -join ";"
    }

    $Issuebody = "[
        {
            `"op`": `"add`",
            `"path`": `"/fields/System.Title`",
            `"value`": `"$(CleanText $Title)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/Microsoft.VSTS.Scheduling.Effort`",
            `"value`": `"$($Effort)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/Microsoft.VSTS.Common.Priority`",
            `"value`": `"$($Priority)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/System.Tags`",
            `"value`": `"$($Tags)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/Microsoft.VSTS.Common.BusinessValue`",
            `"value`": `"$($BusinessValue)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/Microsoft.VSTS.Common.TimeCriticality`",
            `"value`": `"$($TimeCriticality)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/Microsoft.VSTS.Common.Risk`",
            `"value`": `"$($Risk)`"
        },
        {
            `"op`": `"add`",
            `"path`": `"/fields/System.Description`",
            `"value`": $(CleanText $Description)
        },
        {
            `"op`": `"add`",
            `"path`": `"/relations/-`",
            `"value`": $linkedItem
        }        
    ]"

    try {
        
        Write-Host "Adding Work Item: $Title" 
        $postIssueUri = $settings.uriBase + '_apis/wit/workitems/$' + $DevOpsWorkItemType + '?api-version=6.0-preview.2'

        $epicWorkItem = Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $settings.authHeader -Body $Issuebody

    }
    catch {

        Write-Error "Exception while creating work item: $($Issuebody)" 
        Write-Error "$($Error[0].Exception.ToString())"
        #exit

    }
}

Function CleanText {
    param (
        $TextToClean
    )
 
    $outputText = $textToClean -replace "’", "'"

    $outputText = $outputText -replace """root""", "'root'" #aws

    $outputText
}

#Loop through DevOps and add Features for every recommendation in the csv
#Updated function to process import of items for non-Well-Architected assessments
function Add-WorkItemsAdoOther {
    param (
        $settings,
        $assessment
    )

    if ($assessment.recommendations) {
        Write-Host "Fetching existing DevOps Work Items..."

        $existingWorkItems = Get-WorkItemsAdo -settings $settings |
        ForEach-Object {
            @{Title = $_.fields.'System.Title'; Tags = $_.fields.'System.Tags'.Split(';') | ForEach-Object { $_.Trim() } }
        }


        foreach ($item in $assessment.recommendations) {
            try {
                $duplicate = $false

                #Check if exists by ID or Title Name
                if ($null -ne $existingWorkItems) {
                    $duplicateItem = $existingWorkItems | Where-Object { $_.Title -eq $item.'Link-Text' }

                    if ($null -ne $duplicateItem) {
                        if ($duplicateItem.Tags.Contains($item.ReportingCategory)) {
                            $duplicate = $true                            
                        }
                    }
                }

                if ($duplicate -eq $true) {
                    
                    Write-Host "Skipping Duplicate Work Item: $($item.'Link-Text')"
                }
                else {
                    #IF NOT EXISTS
                    #Add Relationship
                    $url = $assessment.reportingCategories[$item.ReportingCategory]
                    $linkedItem = '{"rel": "System.LinkTypes.Hierarchy-Reverse", "url": "EPICURLPLACEHOLDER", "attributes": {"comment": "Making a new link for the dependency"}}'
                    $linkedItem = $linkedItem.Replace("EPICURLPLACEHOLDER", $url)

                    $Priority = "4"
                    $Risk = "1 - High"
                    if ($item.Weight -gt 80) {
                        $Priority = "1"
                        $Risk = "1 - High"
                    }
                    elseif ($item.Weight -gt 60) {
                        $Priority = "2"
                        $Risk = "1 - High"
                    }
                    elseif ($item.Weight -gt 30) {
                        $Priority = "3"
                        $Risk = "2 - Medium"
                    }
                    else {
                        $Priority = "4"
                        $Risk = "3 - Low"
                    }

                    Add-NewIssueToDevOps `
                        -settings $settings `
                        -assessment $assessment `
                        -Title $item.'Link-Text' `
                        -Effort "0" `
                        -Tags $item.ReportingCategory `
                        -Priority $Priority `
                        -BusinessValue $item.Weight `
                        -TimeCriticality $item.Weight `
                        -Risk $Risk `
                        -Description $($item.Description | Out-String | ConvertTo-Json) `
                        -linkedItem $linkedItem
                }
            }
            catch {
                Write-Error "Could not insert item to devops: $($Error[0].Exception.ToString())"
                exit
            }
        }
    }
}

function Add-WorkItemsAdo {
    param (
        $settings,
        $assessment,
        $existingAdoEpics
    )


    if ($assessment.recommendations) {
        Write-Host "Fetching existing DevOps Work Items..."

        $existingWorkItems = Get-WorkItemsAdo -settings $settings |
        ForEach-Object {
            @{Title = $_.fields.'System.Title'; Tags = $_.fields.'System.Tags'.Split(';') | ForEach-Object { $_.Trim() } }
        }

        foreach ($item in $assessment.recommendations) {
            try {
                $duplicate = $false

                #Check if exists by ID or Title Name
                if ($null -ne $existingWorkItems) {
                    $duplicateItem = $existingWorkItems | Where-Object { $_.Title -eq $item.'Link-Text' }

                    if ($null -ne $duplicateItem) {
                        if ($duplicateItem.Tags.Contains($item.Category)) {
                            $duplicate = $true                            
                        }
                    }
                }

                if ($duplicate -eq $true) {
                    
                    Write-Host "Skipping Duplicate Work Item: $($item.'Link-Text')"
                }
                else {
                    # Add Relationship
                    
                    $newReportingCategory = GetMappedReportingCategory -reportingCategrory $item.ReportingCategory # TODO: Check for refactoring (mapping directly in $item.ReportingCategory? outside of this method)
                    $url = $existingAdoEpics[$newReportingCategory]

                    $linkedItem = '{"rel": "System.LinkTypes.Hierarchy-Reverse", "url": "EPICURLPLACEHOLDER", "attributes": {"comment": "Making a new link for the dependency"}}'
                    $linkedItem = $linkedItem.Replace("EPICURLPLACEHOLDER", $url)

                    $Priority = "4"
                    $Risk = "1 - High"
                    if ($item.Weight -gt 80) {
                        $Priority = "1"
                        $Risk = "1 - High"
                    }
                    elseif ($item.Weight -gt 60) {
                        $Priority = "2"
                        $Risk = "1 - High"
                    }
                    elseif ($item.Weight -gt 30) {
                        $Priority = "3"
                        $Risk = "2 - Medium"
                    }
                    else {
                        $Priority = "4"
                        $Risk = "3 - Low"
                    }

                    Add-NewIssueToDevOps `
                        -settings $settings `
                        -assessment $assessment `
                        -Title $item.'Link-Text' `
                        -Effort "0" `
                        -Tags $item.Category `
                        -Priority $Priority `
                        -BusinessValue $item.Weight `
                        -TimeCriticality $item.Weight `
                        -Risk $Risk `
                        -Description $($item.Description | Out-String | ConvertTo-Json) `
                        -linkedItem $linkedItem
                }
            }
            catch {
                Write-Error "Could not insert item to devops: $($Error[0].Exception.ToString())"
                exit
            }
        }
    }
}

#endregion


#region Script Main


# Initialize the dictionaries
$existingAdoEpics = @{}
$reportingCategories = @{}
$categoryMapping = @{}

$adoSettings = Get-AdoSettings

# Assessment type check and import
$isWellArchitected = $assessmentTypeCheck.Contains("Well-Architected")

if ($isWellArchitected) {
    $assessment = Import-Assessment
} else {
    $assessment = Import-AssessmentOther
}


# We ask the end user if they are ready to put data into their ticket system.
Write-Output "Assessment Name: $($assessment.name)" 
Write-Output "URI Base: $($adoSettings.uriBase)"
Write-Output "Number of Recommendations to import: $($assessment.recommendations.Count)" 
Write-Host ""
$confirmation = Read-Host "Ready? [y/n]"

while ($confirmation -ne "y") {
    if ($confirmation -eq 'n') { exit }
    $confirmation = Read-Host "Ready? [y/n]"
}

Write-Output "Processing..."

## Get existing epics from ADO
$existingAdoEpics = Search-EpicsAdo -settings $adoSettings 

#$assessmentCategoryKeys = $assessment.reportingCategories.Keys | Sort-Object
$assessmentCategoryKeys = $assessment.reportingCategories.Keys |
    Sort-Object -Property @{ Expression = { if ($_ -eq 'Defender for Cloud') { 0 } else { 1 } } }, @{ Expression = { $_ } }


foreach ($epicName in $assessmentCategoryKeys) {
    $newEpicName = GetMappedReportingCategory -reportingCategrory $epicName
    if (-not $existingAdoEpics.ContainsKey($newEpicName)) {
            
        $newEpicUrl = Add-EpicAdo -settings $adoSettings -epicName $newEpicName # insert new epic in ADO
        $existingAdoEpics[$newEpicName] = $newEpicUrl # add new epic to the internal structure
    }
}
    
#Insert/update all work items in ADO
Add-WorkItemsAdo -settings $adoSettings -assessment $assessment -existingAdoEpics $existingAdoEpics 


Write-Output ""
Write-Output "Import Complete!"

#endregion