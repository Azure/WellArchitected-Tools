param (
    [string]$pat, 
    [uri]$uri,
    [string]$csv
)

#region Usage

if (!$pat -or !$csv -or !$uri) {
    Write-Host "Example Usage: "
    Write-Host "  PnP-Github.ps1 -pat PAT_FROM_GITHUB -csv ./waf_review.csv -uri https://www.github.com/demo-org/demo-repo"
    Write-Host ""
    exit
}

#endregion

# function Import-Assessment {
#     param (
#         [string]$csv
#     )

    $content = Get-Content $csv
 

    $firstLine = ConvertFrom-Csv $content[0] -Delimiter ',' -Header "Name" | Select-Object -Index 0
    $assessmentName = $firstLine.Name -replace ',' -replace ';'
        
    $tableStart = $content.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context")
    $endStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Hack to change blank reporting category to Azure Advisor
    $devOpsList | 
        Where-Object -Property ReportingCategory -eq "" | 
        ForEach-Object {$_.ReportingCategory = "Azure Advisor"}

    $workingDirectory = (Get-Location).Path
    $recommendationHash = Get-Content "$workingDirectory\WASA.json" | ConvertFrom-Json

    # Get unique list of ReportCategory column

    $reportingCategories = @{}
    $devOpsList | 
        Select-Object -Property ReportingCategory -Unique | 
        Sort-Object  -Property ReportingCategory |
        ForEach-Object { $reportingCategories[$_.ReportingCategory] = "" }

    $assessment = @{
        name = $assessmentName
        reportingCategories = $reportingCategories
        recommendations = $devOpsList
        hash = $recommendationHash
    }

#     return $assessment
# }

#Region Search for existing milestones in Github.

#Hithub Milestones are loosly analogous to epics in ADO. 
#First we look for a milestone that already exists. If it does we move to the next milestone.


function New-GithubMilestone 
{
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][string]$Owner,
        [Parameter(Mandatory=$true)][string]$Repository,
        [Parameter(Mandatory=$true)]$Headers
    )

    $Body = @{
            title  = $Title
            description   = $Description
        } | ConvertTo-Json


        try 
        {
            $AllMilestones = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$owner/$repository/milestones" -Headers $Headers -ContentType "application/json"
            Start-Sleep -Seconds 3
            if($AllMilestones.title -notcontains $Title)
            {
                $NewMilestone = Invoke-RestMethod -Method Post -Uri "https://api.github.com/repos/$owner/$repository/milestones" -Body $Body -Headers $Headers -ContentType "application/json"
                Start-Sleep -Seconds 3
            }
        }
        Catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Output "New-GithubMilestone: $Title $ErrorMessage $FailedItem"
        }
    
}
#endregion

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
        Write-Output "Error while querying Azure DevOps for Epics: " + $Error[0].Exception.ToString()
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
        
        Write-Host "Adding Epic to ADO: $epicName"
        $postIssueUri = $settings.uriBase + "_apis/wit/workitems/$" + "Epic" + "?api-version=5.1"
        $epicWorkItem = Invoke-RestMethod -Uri $postIssueUri -Method POST -ContentType "application/json-patch+json" -Headers $settings.authHeader -Body $body
        $assessment.reportingCategories[$epicName] = $epicWorkItem.url
    } catch {
        Write-Output "Error creating Epic in DevOps: " + $Error[0].Exception.ToString()
        exit
    }
}





Write-Host "Assessment Name:" $assessment.name
Write-Host "URI Base:" $adoSettings.uriBase
Write-Host "Number of Recommendations to import": $assessment.recommendations.Count


















# #region Clean the uncategorized data

# foreach($lineData in $CSVInput)
# {
#     if(!$lineData.ReportingCategory)
#     {
#         $lineData.ReportingCategory = "Uncategorized"
#     }
# }

# #endregion

# function Get-RecommendationsFromContentService
# {
# param(
# [parameter (Mandatory=$true, position=1)]
# [string]$contentservice
# )
#     try
#     {            
#         $ContentServiceResult = Invoke-RestMethod -Method Get -uri $($ContentServiceUri + "$contentservice\") -Headers $ContentServiceHeader
#         foreach($row in $ContentServiceResult)
#         {
#                 $listItem = [pscustomobject]@{
#                     "Assessment" = $row.Assessment;
#                     "ID" = $row.Id;
#                     "Name" = $row.Name;
#                     "WhyConsiderThis" = $row.WhyConsiderThis;
#                     "Context" = $row.Context;
#                     "LearnMore" = $row.LearnMore;
#                     "HowToTroubleshoot" = $row.HowToTroubleshoot;
#                     "SuggestedActions" = $row.SuggestedActions;
#                     "Score" = $row.Score;
#                     "Impact" = $row.Impact;
#                     "Effort" = $row.Effort;
#                     "Probability" = $row.Probability;
#                     "Weight" = $row.Weight;
#                     "FocusArea" = $row.FocusArea;
#                     "FocusAreaId" = $row.FocusAreaId;
#                     "ActionArea" = $row.ActionArea;
#                     "ActionAreaId" = $row.ActionAreaId;
#                 }
#                 if(!$RecommendationHash.Contains($listItem))
#                 {
#                 $RecommendationHash.Add($listItem) | Out-Null
#                 }
#         }
#     }
#     catch{Write-Output "Exception in calling content service for $contentservice : " + $Error[0].Exception.ToString()}
# }

# #ContentService
# #$ContentServiceHeader = @{'Ocp-Apim-Subscription-Key'= ''}
# #$ContentServiceUri = "https://serviceshub-api-prod.azure-api.net/content/contentdefinition/v1.0/"
# #$RecommendationHash = New-Object System.Collections.ArrayList
# #Get-RecommendationsFromContentService -contentservice "ASOCA"
# $RecommendationHash = Get-Content "$workingDirectory\WASA.json" | ConvertFrom-Json

# #Add a new Milestone to GitHub
# function New-GithubMilestone 
# {
#     param(
#         [Parameter(Mandatory=$true)][string]$Title,
#         [Parameter(Mandatory=$true)][string]$Description,
#         [Parameter(Mandatory=$true)][string]$Owner,
#         [Parameter(Mandatory=$true)][string]$Repository,
#         [Parameter(Mandatory=$true)]$Headers
#     )

#     $Body = @{
#             title  = $Title
#             description   = $Description
#         } | ConvertTo-Json


#         try 
#         {
#             $AllMilestones = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$owner/$repository/milestones" -Headers $Headers -ContentType "application/json"
#             Start-Sleep -Seconds 3
#             if($AllMilestones.title -notcontains $Title)
#             {
#                 $NewMilestone = Invoke-RestMethod -Method Post -Uri "https://api.github.com/repos/$owner/$repository/milestones" -Body $Body -Headers $Headers -ContentType "application/json"
#                 Start-Sleep -Seconds 3
#             }
#         }
#         Catch {
#             $ErrorMessage = $_.Exception.Message
#             $FailedItem = $_.Exception.ItemName
#             Write-Output "New-GithubMilestone: $Title $ErrorMessage $FailedItem"
#         }
    
# }

# #Add a new Issue to GitHub
# function New-GithubIssue 
# {
#     param(
#         [Parameter(Mandatory=$true)][string]$Title,
#         [Parameter(Mandatory=$true)][string]$Description,
#         [Parameter(Mandatory=$true)]$Label,
#         [Parameter(Mandatory=$true)][string]$Owner,
#         [Parameter(Mandatory=$true)][string]$Repository,
#         [Parameter(Mandatory=$true)][string]$Milestone,
#         [Parameter(Mandatory=$true)]$Headers
#     )

#     $MilestoneID = $($AllMilestones | Where-Object{$_.Title -eq $Milestone} | Select-Object -First 1).Number

#     $Body = @{
#                 title  = "$Title"
#                 body   = "$Description"
#                 labels = $Label
#                 milestone = "$MilestoneID"
#             } | ConvertTo-Json
#         try 
#         {
#             if($AllIssues.title -notcontains $Title)
#             {
#                 $NewIssue = Invoke-RestMethod -Method Post -Uri "https://api.github.com/repos/$owner/$repository/issues" -Body $Body -Headers $Headers -ContentType "application/json"
#                 Start-Sleep -Seconds 3

#             }
#         }
#         Catch {
#             $ErrorMessage = $_.Exception.Message
#             $FailedItem = $_.Exception.ItemName
#             Write-Output "New-GitHubIssue: $ErrorMessage $FailedItem"
#         }

    
# }

# #region GitHub Management

# Write-Output "Checking for existing categories in Github and adding the missing ones as Milestones"

# foreach($content in $CSVInput)
# {
#     $pillar  = $content.Category
#     if($AllMilestones.title -notcontains $("$pillar - " + $content.ReportingCategory))
#     {
#         $categoryDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq $pillar -and $content.ReportingCategory.Contains($_.Category)}).Description
#         if(!$categoryDescription)
#         {
#             $categoryDescription = "Uncategorized"
#         }
#         New-GithubMilestone -Title $("$pillar - " + $content.ReportingCategory) -Description $categoryDescription -Owner $owner -Repository $repository -Headers $Headers
#         $AllMilestones = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$owner/$repository/milestones" -Headers $Headers -ContentType "application/json"
#     }      
# }

# Write-Output "Attempting Github Import for all Issues"

# foreach($Issue in $CSVInput)
# {   
#     $labels = New-Object System.Collections.ArrayList
#     $labels.Add("$assessmentName")
#     if($Issue.category)
#     {
#         $labels.Add($Issue.category) | Out-Null
#     }

#     if($Issue.ReportingCategory)
#     {
#         $labels.Add($Issue.ReportingCategory) | Out-Null
#     }

#     if($Issue.ReportingsubCategory)
#     {
#         $labels.Add($Issue.ReportingsubCategory) | Out-Null 
#     }  
    
#     $recAdded = $false
#     foreach($recom in $RecommendationHash)
#     {
#         if($recom.Name.Trim('.').Contains($Issue.'Link-Text'.Trim('.')))
#         {
#             $recDescription = "<a href=`"$($Issue.Link)`">$($Issue.'Link-Text')</a>" + "`r`n`r`n" + "<p><b>Why Consider This?</b></p>" + "`r`n`r`n" + $recom.WhyConsiderThis + "`r`n`r`n" + "<p><b>Context</b></p>" + "`r`n`r`n" + $recom.Context + "`r`n`r`n" + "<p><b>Suggested Actions</b></p>" + "`r`n`r`n" + $recom.SuggestedActions + "`r`n`r`n" + "<p><b>Learn More</b></p>" + "`r`n`r`n" + $recom.LearnMore
#             $recDescription = $recDescription -replace ' ',' '
#             $recDescription = $recDescription -replace '“','"' -replace '”','"'
#             New-GithubIssue -Title $Issue.'Link-Text' -Description $recDescription -Label $labels -Owner $owner -Repository $repository -Milestone $($($Issue.category) + " - " + $Issue.ReportingCategory) -Headers $Headers
#             $recAdded = $true
#         }
#     }

#     if(!$recAdded)
#     {
#         $recDescription = "<a href=`"$($Issue.Link)`">$($Issue.'Link-Text')</a>"
#         New-GithubIssue -Title $Issue.'Link-Text' -Description $recDescription -Label $labels -Owner $owner -Repository $repository -Milestone $($($Issue.category) + " - " + $Issue.ReportingCategory) -Headers $Headers
#     }
    
         
# }


# #endregion

# #cleanup
# remove-item $workingDirectory\$reportDate.csv
