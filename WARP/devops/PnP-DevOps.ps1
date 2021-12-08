param (
    [string]$pat, 
    [string]$csv, 
    [uri]$uri
)

if (!$pat -or !$csv -or !$uri) {
    Write-Host "Example Usage: "
    Write-Host "  PnP-DevOps.ps1 -pat PAT_FROM_ADO -csv ./waf_review.csv -uri https://dev.azure.com/demo-org/demo-project"
    Write-Host ""
    exit
}

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)")) }

#Get the Azure Devops project URL and re-create it here.
$UriOrganization = $uri.scheme +"://" + $uri.Host +"/" + $uri.segments[1]

#Grab the project name from the dev-ops url given in the command line.
$projectname = $uri.segments[-1].Trim("/")
$projectname = $projectname +"/"

#Get the working directory from the script
$workingDirectory = (Get-Location).Path

$inputfilename = Split-Path $csv -leaf
$content = Get-Content $csv

$firstLine = ConvertFrom-Csv $content[0] -Delimiter ',' -Header "Name" | Select-Object -Index 0
$assessmentName = $firstLine.Name -replace ',' -replace ';'
    
$ExistingDevopsWI = New-Object System.Collections.ArrayList
$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)")) }

$tableStart = $content.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context")
$EndStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
$tableEnd = $content.IndexOf($EndStringIdentifier) - 1
$DevOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

# Hack to change blank reporting category to Azure Advisor
$DevOpsList | 
    Where-Object -Property ReportingCategory -eq "" | 
    ForEach-Object {$_.ReportingCategory = "Azure Advisor"}

#we ask the end user if they are ready to put data into their ticket system.
Write-Output "This script is using the WAF report:" $inputfilename
Write-Host "This script will insert data into Azure DevOps org:" $UriOrganization.Trim("/")"."
Write-Host "This will insert" $DevOpsList.Length "items into the" $projectname.Trim("/") "project."
Write-Host "We are using the Azure DevOps token that starts with "$pat.substring(0, 5)
$confirmation = Read-Host "Ready? [y/n]"
while($confirmation -ne "y")
{
    if ($confirmation -eq 'n') {exit}
    $confirmation = Read-Host "Ready? [y/n]"
}

# Get unique list of ReportCategory column
$ReportingCategories = $DevOpsList | 
    Select-Object -Property ReportingCategory -Unique | 
    Sort-Object  -Property ReportingCategory

$EpicRelationshipStringBuilder = @'
{"rel": "System.LinkTypes.Hierarchy-Reverse", "url": "EPICURLPLACEHOLDER", "attributes": {"comment": "Making a new link for the dependency"}}
'@

$EpicRelations = @{}
$ReportingCategories | 
    ForEach-Object { $EpicRelations[$_.ReportingCategory] = "" }

$RecommendationHash = Get-Content "$workingDirectory\WASA.json" | ConvertFrom-Json

#Search DevOps for existing Epics for each WAF Category & Create a relationship mapping to link these epics to work items
function Update-EpicsforFocusArea
{

    $body = "{
    `"query`": `"SELECT
        [System.Id],
        [System.WorkItemType],
        [System.Title],
        [System.AssignedTo],
        [System.State],
        [System.Tags]
    FROM workitems
    WHERE
        [System.TeamProject] = @project
        AND [System.WorkItemType] = 'Epic'
        AND [System.State] <> ''`"}"

    $getQueryUri = $UriOrganization + $projectname + "_apis/wit/wiql?api-version=6.0-preview.2"
    $allEpics = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $AzureDevOpsAuthenicationHeader -Body $body

    $ExistingFocusAreas = @{}
    $ReportingCategories | 
        ForEach-Object { $ExistingFocusAreas[$_.ReportingCategory] = $false }

    try {
        #Gather details per devops item
        Write-Output "There are $($allEpics.workItems.Count) Epics in DevOps"

        if($allEpics.workItems.Count -gt 0)
        {
            foreach($epic in $allEpics.workItems.id)
            {
                $getEpicQueryUri = $UriOrganization + $projectname + "_apis/wit/workitems/" + $epic + "?api-version=6.0-preview.2"
                $epicWorkItem = Invoke-RestMethod -Uri $getEpicQueryUri -Method GET -ContentType "application/json" -Headers $AzureDevOpsAuthenicationHeader
                
                $epicTitle = $epicWorkItem.fields.'System.Title'
                if ($ExistingFocusAreas.ContainsKey($epicTitle)) {
                    $ExistingFocusAreas[$epicTitle] = $true;
                    $EpicRelations[$epicTitle] = $EpicRelationshipStringBuilder.Replace("EPICURLPLACEHOLDER",$epicWorkItem.url);
                }
            }
        }

        foreach($key in $ExistingFocusAreas.Keys)
        {
            if (!$ExistingFocusAreas[$key]) {
                Write-Host "Adding Epic: $key"
                $epicWorkItem = Add-EpicInDevOps -FocusArea $key
                $EpicRelations[$key] = $EpicRelationshipStringBuilder.Replace("EPICURLPLACEHOLDER",$epicWorkItem.url);
            }
        }

    } catch {
        Write-Output "Error while querying devops for work items: " + $Error[0].Exception.ToString()
        exit
    }
}

#Create the Epic in DevOps per category/focus area
function Add-EpicInDevOps
{
    param(
        [parameter (Mandatory=$true)]
        [string]$FocusArea
    )

    try {

        $body = "[
            {
                `"op`": `"add`",
                `"path`": `"/fields/System.Title`",
                `"value`": `"$($FocusArea)`"
            }
        ]"

        $postIssueUri = $Uriorganization + $projectname + "_apis/wit/workitems/$" + "Epic" + "?api-version=5.1"
        $epicWorkItem = Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $AzureDevOpsAuthenicationHeader -Body $body
        return $epicWorkItem
             
    } catch {
        Write-Output "Error creating Epic in DevOps: " + $Error[0].Exception.ToString()
        exit
    }
}

#Loop through DevOps and add Features for every recommendation in the csv
function Add-DevOpsList
{
    if($DevOpsList)
    {
        Write-Output "Fetching existing DevOps Work Items"

        Get-AllDevOpsWorkItems

        $AllRecommendations = $ExistingDevopsWI.fields.'System.Title'
        foreach($devopsItem in $DevOpsList)
        {
            try
            {
                #Check if exists by ID or Title Name
                if($null -ne $AllRecommendations -and $AllRecommendations.Contains($devopsItem.'Link-Text'))
                {
                    Write-Host "Skipping Duplicate Work Item: $($devopsItem.'Link-Text')"
                }
                else
                {
                    #IF NOT EXISTS
                    #Add Relationship
                    $linkedItem = $EpicRelations[$devopsItem.ReportingCategory];
                    
                    $Priority = "4"
                    $Risk = "1 - High"
                    if($devopsItem.Weight -gt 80)
                    {
                        $Priority = "1"
                        $Risk = "1 - High"
                    }
                    elseif($devopsItem.Weight -gt 60)
                    {
                        $Priority = "2"
                        $Risk = "1 - High"
                    }
                    elseif($devopsItem.Weight -gt 30)
                    {
                        $Priority = "3"
                        $Risk = "2 - Medium"
                    }
                    else
                    {
                        $Priority = "4"
                        $Risk = "3 - Low"
                    }

                    $recAdded = $false
                    foreach($recom in $RecommendationHash)
                    {
                        if($recom.Name.Trim('.').Contains($devopsItem.'Link-Text'.Trim('.')))
                        {
                            $recDescription = "<a href=`"$($devopsItem.Link)`">$($devopsItem.'Link-Text')</a>" + "`r`n`r`n" + "<p><b>Why Consider This?</b></p>" + "`r`n`r`n" + $recom.WhyConsiderThis + "`r`n`r`n" + "<p><b>Context</b></p>" + "`r`n`r`n" + $recom.Context + "`r`n`r`n" + "<p><b>Suggested Actions</b></p>" + "`r`n`r`n" + $recom.SuggestedActions + "`r`n`r`n" + "<p><b>Learn More</b></p>" + "`r`n`r`n" + $recom.LearnMore
                            $recDescription = $recDescription -replace ' ',' '
                            $recDescription = $recDescription -replace '“','"' -replace '”','"'
                            Add-NewIssueToDevOps -Title $devopsItem.'Link-Text' -Effort "" -Tags $devopsItem.Category -Priority $Priority -BusinessValue $devopsItem.Weight -TimeCriticality $devopsItem.Weight -Risk $Risk -Description $($recDescription | Out-String | ConvertTo-Json) -linkedItem $linkedItem
                            $recAdded = $true
                        }
                    }

                    if(!$recAdded)
                    {
                        $recDescription = "<a href=`"$($devopsItem.Link)`">$($devopsItem.'Link-Text')</a>"
                        Add-NewIssueToDevOps -Title $devopsItem.'Link-Text' -Effort $devopsItem.Weight -Tags $devopsItem.Category -Priority $Priority -BusinessValue $devopsItem.Weight -TimeCriticality $devopsItem.Weight -Risk $Risk -Description $($recDescription | Out-String | ConvertTo-Json) -linkedItem $linkedItem
                    }
                }
            }
            catch
            {
            Write-Output "Could not insert item to devops: " + $Error[0].Exception.ToString()
            }
        }
    }
}

#Retrieve all work items from DevOps
function Get-AllDevOpsWorkItems
{

    #Iterate entire devops issues
    $body = "{
    `"query`": `"Select * From WorkItems Where [Work Item Type] = 'Feature' AND [State] <> 'Closed' AND [State] <> 'Removed' AND [System.TeamProject] = @project order by [Microsoft.VSTS.Common.Priority] asc, [System.CreatedDate] desc`"
    }"

    $getQueryUri = $UriOrganization + $projectname + "_apis/wit/wiql?api-version=6.0-preview.2"
    $AllWI = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $AzureDevOpsAuthenicationHeader -Body $body

    try {
        #Gather details per devops item
        if($AllWI.workItems.Count -gt 0)
        {
            foreach($wi in $AllWI.workItems.id)
            {
                $getWIQueryUri = $UriOrganization + $projectname + "_apis/wit/workitems/" + $wi + "?api-version=6.0-preview.2"
                $workItem = Invoke-RestMethod -Uri $getWIQueryUri -Method GET -ContentType "application/json" -Headers $AzureDevOpsAuthenicationHeader
                $ExistingDevopsWI.Add($workItem) | Out-Null
            }
        }
        else
        {
            Write-Output "There are no work items of type Issue in DevOps yet"
        }
    } catch {
        Write-Output "Error while querying devops for work items: " + $Error[0].Exception.ToString()
    }
}

#Insert Feature into DevOps
function Add-NewIssueToDevOps($Title,$Effort,$Tags,$Priority,$BusinessValue,$TimeCriticality,$Risk,$Description,$linkedItem)
{
   
    if($Title -eq "" -or $null -eq $Title){$Title="NA"}
    if($Effort -eq "" -or $null -eq $Effort){$Effort="0"}
    #if($Tags -eq "" -or $null -eq $Tags){$Tags="NA"}
    if($Priority -eq "" -or $null -eq $Priority){$Priority="4"}
    if($BusinessValue -eq "" -or $null -eq $BusinessValue){$BusinessValue="0"}
    if($TimeCriticality -eq "" -or $null -eq $TimeCriticality){$TimeCriticality="0"}
    if($Risk -eq "" -or $null -eq $Risk){$Risk="3 - Low"}
    if($Description -eq "" -or $null -eq $Description){$Description="NA"}

    
    if($Tags -eq "" -or $null -eq $Tags) {
        $Tags = $assessmentName
    } else {
        $Tags = @($Tags, $assessmentName) -join ";"
    }

    $Issuebody = "[
        {
            `"op`": `"add`",
            `"path`": `"/fields/System.Title`",
            `"value`": `"$($Title)`"
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
            `"value`": $Description
        },
        {
            `"op`": `"add`",
            `"path`": `"/relations/-`",
            `"value`": $linkedItem
        }        
    ]"

    try {
        Write-Host "Adding Work Item: $Title"
        $postIssueUri = $UriOrganization + $projectname + "_apis/wit/workitems/$" + "Feature?api-version=5.1"
        Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $AzureDevOpsAuthenicationHeader -Body $Issuebody

    } catch {

        Write-Output "Exception while creating work item: $($Issuebody)" + $Error[0].Exception.ToString() 
        
    }
}



#region DevOps Management

Write-Output "Checking for existing categories in DevOps and adding the missing ones as Epics"
Update-EpicsforFocusArea

Write-Output "Attempting DevOps Import for all Issues"
Add-DevOpsList

Write-Output ""
Write-Output "Import Complete!"

#endregion
