<#
.SYNOPSIS
    Creates Milestones and Issues in a GitHub repository based on Well-Architected Assessment / Cloud Adoption Security Assessment .csv file.
    
.DESCRIPTION
    Creates Milestones and Issues in a GitHub repository based on Well-Architected Assessment / Cloud Adoption Security Assessment .csv file.

.PARAMETER GithubPersonalAccessToken
    Personal Access Token from Github - find in personal menu (top right), Settings, Developer Settings, Tokens. Token needs Full Access to target Repo.

.PARAMETER GithubrepoUri
    URI of the Github repo
    
.PARAMETER AssessmentCsvPath
    .csv file exported from Well-Architected Assessment / Cloud Adoption Security Assessment

.PARAMETER GithubTagName
    Name of assessment. Note tag cannot be longer than 50 characters. They will be truncated if longer.

.OUTPUTS
    Status message text

.EXAMPLE
    .\PnP-Github -GithubPersonalAccessToken xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx -GithubrepoUri https://github.com/user/repo -GithubTagName WAF -AssessmentCsvPath c:\temp\Azure_Well_Architected_Review_Jan_1_2023_1_00_00_PM.csv
    Adds items from a Well-Architected Assessment .csv export to a Github repository, as Issues with associated Milestones.

.NOTES
    Make sure 'WAF Category Descriptions.csv' is in the same directory as this script. It is used for well-architected assessments to map the old category names to the new category names 
    (Run install-warptools.ps1 if you didn't already)

.LINK

#>

[CmdletBinding()]
param (
    [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$GithubPersonalAccessToken,
    [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][uri]$GithubrepoUri,
    [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$GithubTagName,
    [parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][System.IO.FileInfo]$AssessmentCsvPath
)

$ErrorActionPreference = "break"

# We need to communicate using TLS 1.2 against GitHub.
[Net.ServicePointManager]::SecurityProtocol = 'tls12'
#endregion

#region Get-GithubRateLimit
# We wait at least 1 second between each call per https://docs.github.com/en/rest/guides/best-practices-for-integrators#dealing-with-secondary-rate-limits

function Get-GithubRateLimit {
    param (
        [string]$ratelimit
    )
    Start-Sleep -Seconds 1
    # Write-Output "Ratelimit $ratelimit"

    if ($ratelimit -ge 1 -and $ratelimit -le 2000) {
        Write-Output "Pausing 10 seconds for Github rate-limiting"
        Start-Sleep -Seconds 10
    }
}
#endregion

#region Github-Wait-Timer.
# Wait for secondary ratelimit
# We wait at least 1 second between each call per https://docs.github.com/en/rest/guides/best-practices-for-integrators#dealing-with-secondary-rate-limits

function Github-Wait-Timer {
    param (
        [int32]$Seconds
    )

    $EndTime = [datetime]::UtcNow.AddSeconds($Seconds)
    while (($TimeRemaining = ($EndTime - [datetime]::UtcNow)) -gt 0) {
        Write-Progress -Activity 'Waiting' $seconds -Status 'to let Github rest and allow us to keep working.' -SecondsRemaining $TimeRemaining.TotalSeconds
    }
}
#endregion

#region Get settings for Github
# Github expects to see an authorization token to perform anything interesting. Here we setup the authorization token as a header.
# Example "Authorization: token ghp_xxxxxxxxxxxxxxxxxxxxxx"
function Get-GithubSettings {

    #To reduce the amount of data entry our customers need to do at the command we derive the owner and repository from the URI given.
    $uriBase = $GithubrepoUri.ToString().Trim("/") + "/"

    $owner = $GithubrepoUri.Segments[1].replace('/','')
    $repository = $GithubrepoUri.Segments[2].replace('/','')

    $Headers = @{
        Authorization = 'token ' + $GithubPersonalAccessToken
        }
    $settings = @{
        uriBase = $uriBase
        owner = $owner
        repository = $repository
        pat = $GithubPersonalAccessToken
        Headers = $Headers
    }
    return $settings
}
#endregion

function GetMappedReportingCategory {
    <#
    .DESCRIPTION
    This function takes an old category name as input and returns a new category name. The category name is used as the epic name in Azure DevOps.
    It uses a mapping stored in a hashtable $categoryMapping to find and return the corresponding new category name.  $categoryMapping is build from WAF Category Description.csv
    If the old category name does not exist in the mapping, the function returns the old category name.
    #>    
    param (
        $reportingCategory
    )

    $newReportingCategory = if ($null -ne $categoryMapping -and $categoryMapping.ContainsKey($reportingCategory)) {  
        $categoryMapping[$reportingCategory] # map the old category to the new category # map the old category to the new category
    }
    else {
        $reportingCategory # no mapping found, keep the old category
    }

    return $newReportingCategory
}  


#region function Import-Assessment
# We import the .csv file into memory after making a few housekeeping changes.
function Import-Assessment {

    $workingDirectory = (Get-Location).Path
    [System.IO.FileInfo]$sourceScript = $PSCmdlet.MyInvocation.MyCommand.Source 
    $workingDirectory = $sourceScript.DirectoryName

    try {
        $descriptionsFile = Import-Csv -Path "$workingDirectory\WAF Category Descriptions.csv"
    }
    catch {
        Write-Error -Message "Unable to open $workingDirectory\WAF Category Descriptions.csv"
        exit
    }

    $content = Get-Content $AssessmentCsvPath

    # the table starts at a line of text that looks like the text below and ends with a "--"
    $tableStartPattern = ($content | Select-String "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context" | Select-Object * -First 1)
    $tableStart = ( $tableStartPattern.LineNumber ) - 1
    $endStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Defender for Cloud recommendations do not have a reporting category so we add "Defender for Cloud" as a default to make everything pretty.
    $devOpsList | 
        Where-Object -Property ReportingCategory -eq "" | 
        ForEach-Object {$_.ReportingCategory = "Defender for Cloud"}


    # Get unique list of ReportCategory column
    # Map the existing assessment ReportCategory values to the new categories defined in the 'WAF Category Description.csv' file. The mappings are stored in the $categoryMapping hashtable.
    # These mapped values will then be used as epics and milestones
    $reportingCategories = @{}
    $devOpsList | 
        Select-Object -Property ReportingCategory, Category -Unique | 
        Sort-Object  -Property ReportingCategory |
        ForEach-Object { 

            $currentReportingCategory = $_.ReportingCategory
            $currentPillar = $_.Category
            $categoryTitle = ($descriptionsFile | Where-Object { $_.Pillar -eq $currentPillar -and $_.Category.StartsWith($currentReportingCategory) }).Caption
            if (-not $categoryTitle) {
                $categoryTitle = $currentReportingCategory # Fallback to existing ReportingCategory if no mapping found
            }
            $categoryMapping[$currentReportingCategory] = $categoryTitle
            $reportingCategories[$categoryTitle] = ""       
        }

    # Add Decription 
    $devOpsList | Add-Member -Name Description -MemberType NoteProperty -Value ""

    $devOpsList | 
        ForEach-Object { 
            $_.Description = "<a href=`"$($_.Link)`">$($_.'Link-Text')</a>"
        }

    $githubMilestones = [Ordered]@{}
    $devOpsList | 
        Select-Object -Property Category, ReportingCategory -Unique | 
        ForEach-Object {
            $githubMilestones[(GetMappedReportingCategory -reportingCategory $_.ReportingCategory)] = ""
        }

        $assessment = @{
        name = $GithubTagName
        reportingCategories = $reportingCategories
        recommendations = $devOpsList
        milestones = $githubMilestones
    }
    return $assessment
}
#endregion

#Region function Get-GithubIssues.
function Get-GithubIssues
{
    param (
        $settings
    )
    Write-Output "Fetching existing Github Issues"
    $issuesuri  = "https://api.github.com/repos/" + $settings.owner + "/" + $settings.repository + "/issues?state=open"
    $AllGithubIssues = Invoke-RestMethod $issuesuri -FollowRelLink -MaximumFollowRelLink 10 -Headers $settings.Headers -ResponseHeadersVariable responseHeaders

    $ratelimit = ($responseHeaders.'X-RateLimit-Remaining')
    Write-Output "Rate $ratelimit"
    Get-GithubRateLimit -ratelimit $ratelimit

    if($AllGithubIssues.id.Count -eq 0){
        $AllGithubIssues = @{
            url = "null"
            repository_url = "null"
            labels_url = "null"
            comments_url = "null"
            events_url = "null"
            html_url = "null"
            id = "null"
            node_id = "null"
            number = "null"
            title = "null"
            user = "null"
            labels = "null"
            state = "null"
            locked = "null"
            assignee = "null"
            assignees = "null"
            milestone = "null"
            comments = "null"
            created_at = "null"
            updated_at = "null"
            closed_at = "null"
            author_association = "null"
            active_lock_reason = "null"
            body = "null"
            reactions = "null"
            timeline_url = "null"
            performed_via_github_app = "null"
        }
    }

    return $AllGithubIssues
}
#endregion
# $AllGithubIssues = Get-GithubIssues -settings $settings

#Region function Get-GithubMilestones
function Get-GithubMilestones
{
    param (
        $settings
    )
    Write-Output "Fetching existing Github milestones"
    $milestoneuri = "https://api.github.com/repos/" + $settings.owner + "/" + $settings.repository + "/milestones"

    $AllGithubMilestones  = Invoke-RestMethod $milestoneuri -FollowRelLink -MaximumFollowRelLink 10 -Headers $settings.Headers -ResponseHeadersVariable responseHeaders


    $ratelimit = ($responseHeaders.'X-RateLimit-Remaining')
    Write-Output "Rate $ratelimit"
    Get-GithubRateLimit -ratelimit $ratelimit

    if($AllGithubMilestones.id.Count -eq 0){
        $AllGithubMilestones = @{
            url = "null"
            html_url = "null"
            labels_url = "null"
            id = "null"
            node_id = "null"
            number = "null"
            title = "null"
            description = "null"
            creator = "null"
            open_issues = "null"
            closed_issues = "null"
            state = "null"
            created_at = "null"
            updated_at = "null"
            due_on = "null"
            closed_at = "null"
        }
    }
    return $AllGithubMilestones    
}
# Get-GithubMilestones -settings $settings

#endregion

#Region function Add-MilestoneGithub
#Add a new Milestones to Github
function Add-MilestoneGithub {
    param (
        $settings,
        $milestone,
        $AllMilestones
    )
    $Body = @{
        title = $milestone
        description = ""
    } | ConvertTo-Json
        
    $uri = "https://api.github.com/repos/" + $settings.owner + "/" + $settings.repository + "/milestones"

    try {
        if($AllMilestones.title.Contains($milestone)) {
            # Write-Output " :| Github milestone: $milestone already exists"
        } else {
            $NewMilestone = Invoke-RestMethod -Method Post -Uri $uri -Verbose:$false -Body $Body -Headers $settings.Headers -ContentType "application/json" -ResponseHeadersVariable responseHeaders
            Write-Output " :) We created a new Github milestone: $milestone"
            $ratelimit = ($responseHeaders.'X-RateLimit-Remaining')
            # Write-Output "Rate $ratelimit"
            Get-GithubRateLimit -ratelimit $ratelimit
        }
    }
    Catch {
        $ErrorMessage = $_.Exception.Message
        Write-Output " :( Failure: $milestone $ErrorMessage $responseHeaders"
    }
}

#endregion

#region function Add-GithubIssue
function Add-GithubIssue {
    param (
        $settings,
        $title,
        $bodytext,
        $labels,
        $milestoneid,
        $AllGithubIssues
    )

    if ($AllGithubIssues.title -eq $title) {
        Write-Output "Issue already exists: $title"
    }
    else {
 
        $Body = @{
            title     = $title
            body      = $bodytext
            labels    = $Labels
            milestone = "$MilestoneID"
        } | ConvertTo-Json

        $uri = "https://api.github.com/repos/" + $settings.owner + "/" + $settings.repository + "/issues"
        write-host "Attempting to create a new Github Issue: $title"
        
        try {
            $NewIssue = Invoke-RestMethod -Method Post -Uri $uri -Verbose:$false -Body $Body -Headers $settings.Headers -ContentType "application/json" -ResponseHeadersVariable responseHeaders -MaximumRetryCount 6 -RetryIntervalSec 10
            Write-Output " :) Success"
        
            $ratelimit = ($responseHeaders.'X-RateLimit-Remaining')
            Get-GithubRateLimit -ratelimit $ratelimit
        }
        Catch {
            Write-Output "Response from GitHub: $_.Exception.Message"
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "ReasonPhrase:" $_.Exception.Response.ReasonPhrase
        
            if ($_.Exception.Response.StatusCode.value__ -eq 403) {
                # Try again just for fun!
                try {
                    $NewIssue = Invoke-RestMethod -Method Post -Uri $uri -Verbose:$false -Body $Body -Headers $settings.Headers -ContentType "application/json" -ResponseHeadersVariable responseHeaders -MaximumRetryCount 6 -RetryIntervalSec 10
                    Write-Output " :) Success"
                } Catch {
                    if ($_.Exception.Response.StatusCode.value__ -eq "403") {
                        Github-Wait-Timer -seconds 300
                    }
                }                
            }
            elseif ($_.Exception.Response.StatusCode.value__ -eq 422) {
                Write-Output "Response from GitHub: $_.Exception.Message"
            }
        }
    }
}

#endregion

#region Script Main
$settings = Get-GithubSettings

$categoryMapping = @{}

$assessment = Import-Assessment

#region Ask End User
# We ask the end user if they are ready to put data into their ticket system.
Write-Host "Assessment Name:" $GithubTagName
Write-Host "Repository:" $GithubrepoUri
Write-Host "Number of Recommendations to import": $assessment.recommendations.Count
$confirmation = Read-Host "Ready? [y/n]"
while($confirmation -ne "y")
{
    if ($confirmation -eq 'n') {exit}
    $confirmation = Read-Host "Ready? [y/n]"
}
Write-Host ""
#endregion

#region create new Milestones in Github
# Create new Milestones in Github
# We wait at least 1 second between each call per https://docs.github.com/en/rest/guides/best-practices-for-integrators#dealing-with-secondary-rate-limits
# Search for existing milestones in github before we create new ones.

$AllMilestones = Get-GithubMilestones -settings $settings
Write-Output "Creating Milestones in Github..."
Write-Output ""


$sortedMilestones = $assessment.milestones.Keys |
 Sort-Object @{ Expression = { if ($_ -like '*Defender for Cloud') { 0 } else { 1 } } }, { $_ }

foreach ($milestone in $sortedMilestones) {
    Add-MilestoneGithub -settings $settings -milestone $milestone -allmilestones $AllMilestones
}

Write-Output "All finished creating Milestones in Github..."
Write-Output ""

#endregion

# Region create issues in Github
# Search for existing milestones again in github to reference when we create issues.
# We run this 2 times due to secondary rate limits. These rate limits are undocumented
# We get to run this 3x to get around secondary throttling in github. 
# https://github.com/cli/cli/issues/4801#issuecomment-977747160
# We wait at least 1 second between each call per https://docs.github.com/en/rest/guides/best-practices-for-integrators#dealing-with-secondary-rate-limits

$AllMilestones = Get-GithubMilestones -settings $settings
$AllGithubIssues = Get-GithubIssues -settings $settings

Write-Output "Creating Issues in Github..."
Write-Output ""

# loop through all the assesment items and build the output needed for the function.

foreach($item in $assessment.recommendations){
    $issuetitle=$item.'Link-Text'
    if(!$issuetitle){
        Write-Information 'Issue has no title'
        continue #lel, my continue statement be working, but Write-Information doesn't...
    }

    $bodytext=$item.Description
    $MilestoneName = (GetMappedReportingCategory -reportingCategory $item.ReportingCategory)

    $count = $AllMilestones.Count
    for ($i=0; $i -lt $count; $i++){
        $MilestoneID = ($AllMilestones[$i] | Where-Object{$_.title -eq $MilestoneName}).Number

        if($MilestoneID -is [Int64]){
            break
        }
    }

     # there are some issues with the tag length. Truncate them to 50 characters
    # start gathering labels from the the assesment items 
    $labels = New-Object System.Collections.ArrayList
    $charLimit = 50

    if ($GithubTagName -and $labels -notcontains $GithubTagName) {$labels.Add($GithubTagName.Substring(0, [Math]::Min($GithubTagName.Length, $charLimit))) | Out-Null}
    if ($item.Category -and $labels -notcontains $item.Category) { $labels.Add($item.Category.Substring(0, [Math]::Min($item.Category.Length, $charLimit))) | Out-Null }
    if ($item.ReportingCategory -and $labels -notcontains $item.ReportingCategory) {$labels.Add($item.ReportingCategory.Substring(0, [Math]::Min($item.ReportingCategory.Length, $charLimit))) | Out-Null}
  

    # put all info into github
    Add-GithubIssue -settings $settings -title $issuetitle -bodytext $bodytext -labels $labels -milestoneid $milestoneid -AllGithubIssues $AllGithubIssues

}

Write-Output "All finished creating Issues in Github..."
Write-Output ""

#endregion
#endregion
