#region Parameters
# Command line paramaters required.
# pat = Personal Access Token from Github or ADO
# URI = the URL for the ADO Project or the Github repo
# CSV = The exported CSV file from the WAF Assesment

param (
    [string]$pat,
    [uri]$uri,
    [string]$csv,
    [string]$name = "WAF-Review"
)

#region Usage

if (!$pat -or !$csv -or !$uri -or !$name) {
    Write-Output "Example Usage: "
    Write-Output "  PnP-DevOps.ps1 -pat PAT_FROM_ADO -csv ./waf_review.csv -uri https://dev.azure.com/demo-org/demo-project -name WAF-Assessment-x"
    Write-Output ""
    exit
}

#endregion

#region Functions

# Get settings for either Azure DevOps
function Get-AdoSettings {
    param (
        [string]$pat, 
        [uri]$uri
    )

    $authHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)")) }
    
    $uriBase = $uri.ToString().Trim("/") + "/"

    $settings = @{
        authHeader = $authHeader
        uriBase = $uriBase
    }

    return $settings
}

function Import-Assessment {
    param (
        [string]$csv,
        [string]$name
    )

    $content = Get-Content $csv
        
    $tableStart = $content.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context")
    $endStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Hack to change blank reporting category to Azure Advisor
    $devOpsList | 
        Where-Object -Property ReportingCategory -eq "" | 
        ForEach-Object {$_.ReportingCategory = "Azure Advisor"}

    # get the WASA,json file in an xplat form.
    $workingDirectory = (Get-Location).Path
    $WASAFile = Join-Path -Path $workingDirectory -ChildPath 'WAF.json'
    $recommendationDetail = Get-Content $WASAFile | ConvertFrom-Json

    # Get unique list of ReportCategory column
    $reportingCategories = @{}
    $devOpsList | 
        Select-Object -Property ReportingCategory -Unique | 
        Sort-Object  -Property ReportingCategory |
        ForEach-Object { 
            $reportingCategories[$_.ReportingCategory] = ""       
        }

    # Add Decription and augment it using WAF.json data (if exists)
    $devOpsList | Add-Member -Name Description -MemberType NoteProperty -Value ""

    $devOpsList | 
        ForEach-Object { 

            $_.Description = "<a href=`"$($_.Link)`">$($_.'Link-Text')</a>"

            foreach($detail in $recommendationDetail)
            {
                $detailName = $detail.Name.Trim('.')
                $linkText = $_.'Link-Text'.Trim('.')

                if(($detailName.Contains($linkText)))
                {
                    $recDescription = "<a href=`"$($_.Link)`">$($_.'Link-Text')</a>" + "`r`n`r`n" `
                    + "<p><b>Why Consider This?</b></p>" + "`r`n`r`n" + $detail.WhyConsiderThis + "`r`n`r`n" `
                    + "<p><b>Context</b></p>" + "`r`n`r`n" + $detail.Context + "`r`n`r`n" `
                    + "<p><b>Suggested Actions</b></p>" + "`r`n`r`n" + $detail.SuggestedActions + "`r`n`r`n" `
                    + "<p><b>Learn More</b></p>" + "`r`n`r`n" + $detail.LearnMore
                    
                    $recDescription = $recDescription -replace ' ',' '
                    $recDescription = $recDescription -replace '“','"' -replace '”','"'

                    $_.Description = $recDescription

                    break
                }
            }           
        }

    $assessment = @{
        name = $name
        reportingCategories = $reportingCategories
        recommendations = $devOpsList
    }

    return $assessment
}

function Search-EpicsAdo {
    param (
        $settings,
        $assessment
    )

    $body = "{
        `"query`": `"SELECT
            [System.Id]
        FROM workitems
        WHERE
            [System.TeamProject] = @project
            AND [System.WorkItemType] = 'Epic'
            AND [System.State] <> ''`"}"
     
    try {
        $getQueryUri = $settings.uriBase + "_apis/wit/wiql?api-version=6.0-preview.2"
        $results = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $settings.authHeader -Body $body
        if($results.workItems.Count -gt 0)
        {
            foreach($epic in $results.workItems.id)
            {
                $getEpicQueryUri = $settings.uriBase + "_apis/wit/workitems/" + $epic + "?api-version=6.0-preview.2"
                $epicWorkItem = Invoke-RestMethod -Uri $getEpicQueryUri -Method GET -ContentType "application/json" -Headers $settings.authHeader
                $epicName = $epicWorkItem.fields.'System.Title'
                if ($assessment.reportingCategories.ContainsKey($epicName)) {
                    $assessment.reportingCategories[$epicName] = $epicWorkItem.url
                }
            }
        }
    } catch {
        Write-Error "Error while querying Azure DevOps for Epics: $($Error[0].Exception.ToString())"
        exit
    }     
}

#Create the Epic in DevOps per category/focus area
function Add-EpicAdo
{
    param (
        $settings,
        $assessment,
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
        
        Write-Output "Adding Epic to ADO: $epicName"
        $postIssueUri = $settings.uriBase + "_apis/wit/workitems/$" + "Epic" + "?api-version=5.1"
        $epicWorkItem = Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $settings.authHeader -Body $body
        write-debug "    Output: $epicworkitem"
        $assessment.reportingCategories[$epicName] = $epicWorkItem.url
    } catch {
        Write-Error "Error creating Epic in DevOps: $($Error[0].Exception.ToString())"
        exit
    }
}

#Retrieve all work items from DevOps
function Get-WorkItemsAdo
{
    param (
        $settings
    )

    #Iterate entire devops issues
    $body = "{
    `"query`": `"Select * From WorkItems Where [Work Item Type] = 'Feature' AND [State] <> 'Closed' AND [State] <> 'Removed' AND [System.TeamProject] = @project order by [Microsoft.VSTS.Common.Priority] asc, [System.CreatedDate] desc`"
    }"

    $getQueryUri = $settings.uriBase + "_apis/wit/wiql?api-version=6.0-preview.2"
    $results = Invoke-RestMethod -Uri $getQueryUri -Method POST -ContentType "application/json" -Headers $settings.authHeader -Body $body


    $workItemsAdo = @()
    try {
        #Gather details per devops item
        if($results.workItems.Count -gt 0)
        {
            foreach($wi in $results.workItems.id)
            {
                $getWIQueryUri = $settings.uriBase + "_apis/wit/workitems/" + $wi + "?api-version=6.0-preview.2"
                $workItem = Invoke-RestMethod -Uri $getWIQueryUri -Method GET -ContentType "application/json" -Headers $settings.authHeader
                $workItemsAdo += $workItem
            }
        }
        else
        {
            Write-Verbose "There are no work items of type Issue in DevOps yet"
        }
    } catch {
        Write-Error "Error while querying devops for work items: $($Error[0].Exception.ToString())"
    }

    return $workItemsAdo
}

#Insert Feature into DevOps
function Add-NewIssueToDevOps
{
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
   
    if($Title -eq "" -or $null -eq $Title){$Title="NA"}
    if($Effort -eq "" -or $null -eq $Effort){$Effort="0"}
    #if($Tags -eq "" -or $null -eq $Tags){$Tags="NA"}
    if($Priority -eq "" -or $null -eq $Priority){$Priority="4"}
    if($BusinessValue -eq "" -or $null -eq $BusinessValue){$BusinessValue="0"}
    if($TimeCriticality -eq "" -or $null -eq $TimeCriticality){$TimeCriticality="0"}
    if($Risk -eq "" -or $null -eq $Risk){$Risk="3 - Low"}
    if($Description -eq "" -or $null -eq $Description){$Description="NA"}

    
    if($Tags -eq "" -or $null -eq $Tags) {
        $Tags = $assessment.name
    } else {
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
        $postIssueUri = $settings.uriBase + "_apis/wit/workitems/$" + "Feature?api-version=5.1"
        $result = Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $settings.authHeader -Body $Issuebody

    } catch {

        Write-Error "Exception while creating work item: $($Issuebody)" 
        Write-Error "$($Error[0].Exception.ToString())"
        #exit

    }
}

Function CleanText{
param (
    $TextToClean
)
 
        $outputText = $textToClean -replace     "’","'"

        $outputText = $outputText -replace     """root""", "'root'" #aws

        $outputText
}

#Loop through DevOps and add Features for every recommendation in the csv
function Add-WorkItemsAdo
{
    param (
        $settings,
        $assessment
    )

    if($assessment.recommendations)
    {
        Write-Host "Fetching existing DevOps Work Items"

        $existingWorkItems = Get-WorkItemsAdo -settings $settings |
            ForEach-Object {
                @{Title = $_.fields.'System.Title'; Tags = $_.fields.'System.Tags'.Split(';')}
            }

        foreach($item in $assessment.recommendations)
        {
            try
            {
                $duplicate = $false

                #Check if exists by ID or Title Name
                if($null -ne $existingWorkItems)
                {
                    $duplicateItem = $existingWorkItems | Where-Object {$_.Title -eq $item.'Link-Text'}

                    if ($null -ne $duplicateItem) {
                        if ($duplicateItem.Tags.Contains($item.Category)) {
                            $duplicate = $true                            
                        }
                    }
                }

                if ($duplicate -eq $true)
                {
                    
                    Write-Host "Skipping Duplicate Work Item: $($item.'Link-Text')"
                }
                else
                {
                    #IF NOT EXISTS
                    #Add Relationship
                    $url = $assessment.reportingCategories[$item.ReportingCategory]
                    $linkedItem = '{"rel": "System.LinkTypes.Hierarchy-Reverse", "url": "EPICURLPLACEHOLDER", "attributes": {"comment": "Making a new link for the dependency"}}'
                    $linkedItem = $linkedItem.Replace("EPICURLPLACEHOLDER", $url)

                    $Priority = "4"
                    $Risk = "1 - High"
                    if($item.Weight -gt 80)
                    {
                        $Priority = "1"
                        $Risk = "1 - High"
                    }
                    elseif($item.Weight -gt 60)
                    {
                        $Priority = "2"
                        $Risk = "1 - High"
                    }
                    elseif($item.Weight -gt 30)
                    {
                        $Priority = "3"
                        $Risk = "2 - Medium"
                    }
                    else
                    {
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
            catch
            {
                Write-Error "Could not insert item to devops: $($Error[0].Exception.ToString())"
                exit
            }
        }
    }
}

#endregion


#region Script Main

$adoSettings = Get-AdoSettings -pat $pat -uri $uri

$assessment = Import-Assessment -csv $csv -name $name

# We ask the end user if they are ready to put data into their ticket system.
Write-Output "Assessment Name: $($assessment.name)" 
Write-Output "URI Base: $($adoSettings.uriBase)"
Write-Output "Number of Recommendations to import: $($assessment.recommendations.Count)" 
$confirmation = Read-Host "Ready? [y/n]"
while($confirmation -ne "y")
{
    if ($confirmation -eq 'n') {exit}
    $confirmation = Read-Host "Ready? [y/n]"
}

# Search for existing Epics in ADO
Search-EpicsAdo -settings $adoSettings -assessment $assessment

# Create new Epics in ADO
$newEpics = $assessment.reportingCategories.GetEnumerator() | 
    Where-Object { $_.Value -eq "" } 

if ($newEpics.Count -gt 0) {
    $newEpics.Key | 
        ForEach-Object {
            Add-EpicAdo -settings $adoSettings -assessment $assessment -epicName $_
        }
}

Write-Output "Attempting DevOps Import for all Issues"
Add-WorkItemsAdo -settings $adoSettings -assessment $assessment

Write-Output ""
Write-Output "Import Complete!"

#endregion
