#region Parameters
# Command line paramaters required.
# pat = Personal Access Token from Github or ADO
# URI = the URL for the ADO Project or the Github repo
# CSV = The exported CSV file from the WAF Assesment

param (
    [string]$pat, 
    [uri]$uri,
    [string]$csv,
    [string]$name
)

# We need to communicate using TLS 1.2 against GitHub.
[Net.ServicePointManager]::SecurityProtocol = 'tls12'
#endregion

#region Usage
# How to use the script.
if (!$pat -or !$csv -or !$uri -or !$name) {
    Write-Host "Example Usage: "
    Write-Host "  PnP-Github.ps1 -pat PAT_FROM_GITHUB -csv ./waf_review.csv -uri https://dev.github.com/demo-org/demo-repo" -name "WAF-Assessment-x"
    Write-Host ""
    exit
}
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
        sleep 10
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
# Example "Authorization: token ghp_16C7e42F292c6912E7710c838347Ae178B4a"
function Get-GithubSettings {
     param (
         [string]$pat, 
         [uri]$uri
     )

    #To reduce the amount of data entry our customers need to do at the command we derive the owner and repository from the URI given.
    $uriBase = $uri.ToString().Trim("/") + "/"

    $owner = $uri.Segments[1].replace('/','')
    $repository = $uri.Segments[2].replace('/','')

    $Headers = @{
        Authorization='token '+$pat
        }
    $settings = @{
        uriBase = $uriBase
        owner = $owner
        repository = $repository
        pat = $pat
        Headers = $Headers
    }
    return $settings
}
#endregion
# $settings = Get-GithubSettings -pat $pat -uri $uri

#region function Import-Assessment
# We import the .csv file into memory after making a few housekeeping changes.
function Import-Assessment {
    param (
        [string]$csv
    )
    $content = Get-Content $csv
    $firstLine = ConvertFrom-Csv $content[0] -Delimiter ',' -Header "Name" | Select-Object -Index 0
 
    # the table starts at a line of text that looks like the text below and ends with a "--"
    $tableStart = $content.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context")
    $endStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
    $tableEnd = $content.IndexOf($endStringIdentifier) - 1
    $devOpsList = ConvertFrom-Csv $content[$tableStart..$tableEnd] -Delimiter ','

    # Azure Advisor recommendations do not have a reporting category so we add "Azure Advisor" as a default to make everything pretty.
    $devOpsList | 
        Where-Object -Property ReportingCategory -eq "" | 
        ForEach-Object {$_.ReportingCategory = "Azure Advisor"}

    # We need to know what WAF.json does and where to get a fresh copy.
    # get the WASA,json file in an xplat form.
    $workingDirectory = (Get-Location).Path
    $WASAFile = Join-Path -Path $workingDirectory -ChildPath 'WAF.json'
    $recommendationHash = Get-Content $WASAFile | ConvertFrom-Json
    
    # Get unique list of ReportCategory column
    # we will use these values as epics and milestones
    $reportingCategories = @{}
    $devOpsList | 
        Select-Object -Property ReportingCategory -Unique | 
        ForEach-Object {
            $reportingCategories[$_.ReportingCategory] = ""
        }

    $githubMilestones = [Ordered]@{}
    $devOpsList | 
        Select-Object -Property Category, ReportingCategory -Unique | 
        ForEach-Object {
            $githubMilestones[$_.Category + " - " + $_.ReportingCategory] = ""
        }

        $assessment = @{
        reportingCategories = $reportingCategories
        recommendations = $devOpsList
        hash = $recommendationHash
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
        Write-Output " :( There was an error creating new Github milestone: $milestone"
        Write-Output "    The $ErrorMessage"
        Write-Output $responseHeaders
    }
}

#endregion

#region function Create-GithubIssue
function Create-GithubIssue {
    param (
        $settings,
        $title,
        $bodytext,
        $labels,
        $milestoneid,
        $AllGithubIssues
    )

    if($AllGithubIssues.title -eq $title) {
        Write-Output "Yes exist: $title"
    } else {
 
        $Body = @{
            title  = $title
            body   = $issuebodytext
            labels = $Labels
            # milestone = "$MilestoneID"
        } | ConvertTo-Json

        $uri = "https://api.github.com/repos/" + $settings.owner + "/" + $settings.repository + "/issues"
        write-host "Attempting to create a new Github Issue: $issuetitle"
        
        try {
            $NewIssue = Invoke-RestMethod -Method Post -Uri $uri -Verbose:$false -Body $Body -Headers $settings.Headers -ContentType "application/json" -ResponseHeadersVariable responseHeaders -MaximumRetryCount 6 -RetryIntervalSec 10
            Write-Output " :)       We created a new Github Issue: $issuetitle"
    
            $ratelimit = ($responseHeaders.'X-RateLimit-Remaining')
            Get-GithubRateLimit -ratelimit $ratelimit
        } Catch {
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
            Write-Host "ReasonPhrase:" $_.Exception.Response.ReasonPhrase
            # Write-Host "All:" $_.Exception.Response
            if ($_.Exception.Response.StatusCode.value__ -eq "403") {
                Github-Wait-Timer -seconds 300
                # Try again just for fun!
                try {
                    $NewIssue = Invoke-RestMethod -Method Post -Uri $uri -Verbose:$false -Body $Body -Headers $settings.Headers -ContentType "application/json" -ResponseHeadersVariable responseHeaders -MaximumRetryCount 6 -RetryIntervalSec 10
                    Write-Output " :)       We created a new Github Issue: $issuetitle"
                } Catch {
                    if ($_.Exception.Response.StatusCode.value__ -eq "403") {
                        Github-Wait-Timer -seconds 300
                    }
                }
            } elseif ($_.Exception.Response.StatusCode.value__ -eq "422") {
                Write-Output " :|       This may be a duplicate Issue: $issuetitle"
            }
        }    
    }
}

#endregion

#region Script Main

$settings = Get-GithubSettings -pat $pat -uri $uri
$assessment = Import-Assessment -csv $csv

#region Ask End User
# We ask the end user if they are ready to put data into their ticket system.
Write-Host "Assessment Name:" $name
Write-Host "Repository:" $uri
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

$assessment.milestones.GetEnumerator() | ForEach-Object{
    Add-MilestoneGithub -settings $settings -milestone $_.key -allmilestones $AllMilestones
}
Write-Output "All finished creating Milestones in Github..."
Write-Output ""

#endregion

#Region create issues in Github
# Search for existing milestones again in github to reference when we create issues.
# We run this 2 times due to secondary rate limits. These rate limits are undocumented
# We get to run this 3x to get around secondary throttling in github. 
# https://github.com/cli/cli/issues/4801#issuecomment-977747160
# We wait at least 1 second between each call per https://docs.github.com/en/rest/guides/best-practices-for-integrators#dealing-with-secondary-rate-limits

$AllMilestones = Get-GithubMilestones -settings $settings
$AllGithubIssues = Get-GithubIssues -settings $settings

Write-Output "Creating Issues in Github... (take 1)"
Write-Output ""

# loop through all the assesment items and build the output needed for the function.

foreach($item in $assessment.recommendations){
    $issuetitle=$item.'Link-Text'
    $body="<a href=`"$($item.Link)`">$($issuetitle)</a>`r`n`r`n"
    $MilestoneName=($item.category + " - " + $item.ReportingCategory)

    foreach($WASA in $assessment.hash){
        if ($WASA.Name -eq $issuetitle){
            $WASAbody="`r`n`r`n<p><b>Why Consider This?</b></p>" + "`r`n" + $WASA.WhyConsiderThis + "`r`n`r`n"
            $WASAbody+="<p><b>Context</b></p>" + "`r`n" + $WASA.Context + "`r`n`r`n"
            $WASAbody+="<p><b>Suggested Actions</b></p>" + "`r`n" + $WASA.SuggestedActions + "`r`n`r`n"
            $WASAbody+="<p><b>Learn More</b></p>" + "`r`n" + $WASA.LearnMore
            $body+=$WASAbody
        }
    }

    # start gathering labels from the the assesment items and the WAF.json
    $labels = New-Object System.Collections.ArrayList
    $labels.Add("WARP-Import $name") | Out-Null
    if($item.category){
        $labels.Add($item.Category) | Out-Null
    }
    if($item.ReportingCategory){
        $labels.Add($item.ReportingCategory) | Out-Null
    }
    if($item.ReportingSubcategory){
        $labels.Add($item.ReportingSubcategory) | Out-Null
    }
    if($WASA.FocusArea){
        $labels.Add($WASA.FocusArea) | Out-Null
    }
    if($WASA.ActionArea){
        $labels.Add($WASA.ActionArea) | Out-Null
    }

    # toss all info 
    Create-GithubIssue -settings $settings -title $issuetitle -bodytext $bodytext -labels $labels -milestoneid $milestoneid -AllGithubIssues $AllGithubIssues

}


Write-Output "All finished creating Issues in Github..."
Write-Output ""

#endregion
#endregion