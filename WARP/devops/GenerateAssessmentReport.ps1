#requires -Version 7
<#
.SYNOPSIS
    Takes output from the Well-Architected Review Assessment website and produces a PowerPoint presentation incorporating the findings.
    Also support the Cloud Adoption Security Review Assessment.
    https://learn.microsoft.com/en-us/assessments/azure-architecture-review/
    
.DESCRIPTION
    The Well-Architected Review site provides a self-Assessment tool. This might be used for self-assessment, or run as part of an Assessment performed by Microsoft to help recommend improvements.

.PARAMETER ContentFile
    Exported CSV file from the Well-Architected Review file. Supports relative paths.

.PARAMETER CloudAdoption
    If set, indicates the Cloud Adoption Security Review format should be used. If not set, Well-Architected is assumed.
    
.PARAMETER MinimumReportLevel
    The level above which a finding is considered high severity, By convention, scores up to 32 are low, 65 medium, and 66+ high.

.PARAMETER ShowTop
    How many recommendations to try to fit on a slide. 8 is default.

.PARAMETER NoSmoke
    This parameter doesn't exist, and whomever told you about it was pulling your leg.

.INPUTS
    ContentFile should be a CSV-formatted Well-Architected Assessment export

.OUTPUTS
    PowerPoint Presentation - WAF-Review-2023-16-1500.pptx
    CSV artifact suitable for use with the DevOps/GitHub import scripts

.EXAMPLE
    .\generateAssessmentReport.ps1  
    
    If no -ContentFile is specified, a file browser dialog box will be shown and the file may be selected.
    
    Generates a PPTX report from a Well-Architected Review site exported CSV.

.EXAMPLE
    .\generateAssessmentReport.ps1 -ContentFile .\Cloud_Adoption_Security_Review_Sample.csv -ShowTop 9
    
    Generates a PPTX report from a CASR CSV
    
    Tries to include the top 9 results of any category (which probably won't fit by default, so plan to reformat things)
        
.EXAMPLE
    .\generateAssessmentReport.ps1 -ContentFile .\Cloud_Adoption_Security_Review_Sample.csv -CloudAdoption    
    
    If the title doesn't identify the report type correctly, you can force the decision with -CloudAdoption

.NOTES
    PowerPoint needs to be installed to create a PPTX.
    The CSV output is filtered when using WAF to work around some data issues - only nominated pillar findings will be processed.
    The Assessment type is attemptedly guessed from the title on the input CSV. If it can't be guessed, WAF is assumed.

.LINK
    https://github.com/Azure/WellArchitected-Tools/
#>
[CmdletBinding()]
param (
    # Indicates CSV file for input
    [Parameter()][string]
    $ContentFile ,

    # Minimum level for inclusion in summary (defaults to High)
    [Parameter()][int]
    $MinimumReportLevel = 65 ,

    # Show Top N Recommendations Per Slide (default 8)
    [Parameter()][int]
    $ShowTop = 6 ,

    [Parameter()]
    [switch] $CloudAdoption

)
<# Instructions to use this script

From the directory in which you've downloaded the scripts, templates and csv file from the PnP survey (mycontent.csv)

For a WAF Assessment report:

    .\GenerateAssessmentReport.ps1 -ContentFile .\mycontent.csv

For a CASR report:

    .\GenerateAssessmentReport.ps1 -ContentFile .\mycontent.csv -CloudAdoption

Ensure the powerpoint template file and the Category Descriptions file exist in the paths shown below before attempting to run this script

Once the script is run, close the powershell window and a timestamped PowerPoint report and a subset csv file will be created on the working directory

Use these reports to represent and edit your findings for the WAF Engagement

Known issues 
  a. Pillar scores may not reflect accurately if the ordering in the csv is jumbled. Please adjust lines 41-53 in case the score representations for the pillars are not accurate
  b. If the hyperlinks are not being published accurately, ensure that the csv file doesnt have any multi-sentence recommendations under Link-Text field

#>
#Get the working directory from the script
$workingDirectory = (Get-Location).Path

Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.Title = "Select Well-Architected Review file export"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

Function FindIndexBeginningWith($stringset, $searchterm) {
    $i = 0
    foreach ($line in $stringset) {
        if ($line.StartsWith($searchterm)) {
            return $i
        }
        $i++
    }
    return false
}

if ([String]::IsNullOrEmpty($ContentFile)) {
    $inputfile = Get-FileName $workingDirectory
}
else {
    if (!(Resolve-Path $ContentFile)) {
        $inputfile = Get-FileName $workingDirectory
    }
    else {
        $inputFile = $ContentFile
    }
}
# validate our file is OK
try {
    $content = Get-Content $inputfile
}
catch {
    Write-Error -Message "Unable to open selected Content file."
    exit
}

$assessmentTypeCheck = ""

#$inputfilename = Split-Path $inputfile -leaf
$assessmentTypeCheck = ($content | Select-Object -First 1)

# initial fix for wrong report type selection
if ($assessmentTypeCheck.contains("Cloud Adoption") ) {
    write-host "Detected Cloud Adoption Security Review from CSV title row"
    $CloudAdoption = $true
}
if (!$CloudAdoption) {
    write-host "Well-Architected Review selected - use -CloudAdoption switch if incorrect."
    $assessmentTypeCheck = "Well-Architected"
}
$reportDate = Get-Date -Format "yyyy-MM-dd-HHmm"
$localReportDate = Get-Date -Format g

#region establish pillars and score from top of report
$overallScore = ""
$costScore = ""
$operationsScore = ""
$performanceScore = ""
$reliabilityScore = ""
$securityScore = ""
$overallScoreDescription = ""

$filteredPillars = @()

if ($assessmentTypeCheck.contains("Well-Architected")) {
    for ($i = 3; $i -le 8; $i++) {
        #write-Debug "Line: $($Content[$i])"
        if ($Content[$i].Contains("overall")) {
            $overallScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        }
        if ($Content[$i].Contains("Cost Optimization")) {
            $costScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
            $filteredPillars += "Cost Optimization"
        }
        if ($Content[$i].Contains("Reliability")) {
            $reliabilityScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
            $filteredPillars += "Reliability"
        }
        if ($Content[$i].Contains("Operational Excellence")) {
            $operationsScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
            $filteredPillars += "Operational Excellence"
        }
        if ($Content[$i].Contains("Performance Efficiency")) {
            $performanceScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
            $filteredPillars += "Performance Efficiency"
        }
        if ($Content[$i].Contains("Security")) {
            $securityScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
            $filteredPillars += "Security"
        }
        if ($Content[$i].Equals(",,,,,")) {
            #End early if not all pillars assessed
            Break
        }
    }
}
else {
    $i = 3
    if ($Content[$i].Contains("overall")) {
        $overallScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $overallScoreDescription = $Content[$i].Split(',')[1]
    }
}
#endregion

if ($assessmentTypeCheck.contains("Well-Architected")) {
    Write-host "Producing Well Architected report from $inputFile"
    $templatePresentation = "$workingDirectory\PnP_PowerPointReport_Template.pptx"
    $title = "Well-Architected [pillar] Assessment" # Don't edit this - it's used when multiple Pillars are included.
    try {
        $tableStart = FindIndexBeginningWith $content "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context"
    }
    catch{
        Write-host "That appears not to be a content file. Please use only content from the Well-Architected Assessment site."
    }
    try{
        #Write-Debug "Tablestart: $tablestart"
        $EndStringIdentifier = $content | Where-Object { $_.Contains("--,,") } | Select-Object -Unique -First 1
        #Write-Debug "EndStringIdentifier: $EndStringIdentifier"
        $tableEnd = $content.IndexOf($EndStringIdentifier) - 1
        #Write-Debug "Tableend: $tableend"
        $csv = $content[$tableStart..$tableEnd] | Out-File  "$workingDirectory\$reportDate.csv"
        $importdata = Import-Csv -Path "$workingDirectory\$reportDate.csv"

        #region Clean the uncategorized data
        #if ($importdata.PSobject.Properties.Name -contains "ReportingCategory") {
        
            foreach ($lineData in $importdata) {
                
                if (!$lineData.ReportingCategory) {
                    $lineData.ReportingCategory = "Uncategorized"
                }
            }
        #}
        #endregion
        $data = $importdata | where {$_.Category -in $filteredPillars}
        $data | Export-Csv -UseQuotes AsNeeded "$workingDirectory\$reportDate.csv" 
        $data | % { $_.Weight = [int]$_.Weight }
        $pillars = $data.Category | Select-Object -Unique
        Write-host -Debug $pillars
    }
    catch {
        Write-Host "Unable to parse the content file."
        Write-Host "Please ensure all input files are in the correct format and aren't open in Excel or another editor which locks the file."
        Write-Host "--"
        Write-Host $_
        exit
    }
}
else {
    Write-host "Producing Cloud Adoption Security Review report..."
    $templatePresentation = "$workingDirectory\PnP_PowerPointReport_Template - CloudAdoption.pptx"
    $title = "Cloud Adoption Security Review"
    try {
        $tableStart = FindIndexBeginningWith $content "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context,CompleteY/N,Note"
        #Write-Debug "Tablestart: $tablestart"
        $EndStringIdentifier = $content | Where-Object { $_.Contains("--,,") } | Select-Object -Unique -First 1
        #Write-Debug "EndStringIdentifier: $EndStringIdentifier"
        $tableEnd = $content.IndexOf($EndStringIdentifier) - 1
        #Write-Debug "Tableend: $tableend"
        $csv = $content[$tableStart..$tableEnd] | Out-File  "$workingDirectory\$reportDate.csv"
        $data = Import-Csv -Path "$workingDirectory\$reportDate.csv"
        $data | % { $_.Weight = [int]$_.Weight }
        #$pillars = $data.Category | Select-Object -Unique
    }
    catch {
        Write-Host "Unable to parse the content file."
        Write-Host "Please ensure all input files are in the correct format and aren't open in Excel or another editor which locks the file."
        Write-Host "--"
        Write-Host $_
        exit
    }
}

if ($assessmentTypeCheck.contains("Well-Architected")) {
    try {
        $descriptionsFile = Import-Csv "$workingDirectory\WAF Category Descriptions.csv"
    }
    catch {
        Write-Error -Message "Unable to open $($workingDirectory)\WAF Category Descriptions.csv"
        exit
    }
}
else {
    try {
        $descriptionsFile = Import-Csv "$workingDirectory\CAF Category Descriptions.csv"
    }
    catch {
        Write-Error -Message "Unable to open $($workingDirectory)\CAF Category Descriptions.csv"
        exit
    }
}


#endregion


#region CSV Calculations
$otherAssessmentDescription = ($descriptionsFile | Where-Object { $_.Category -eq "Survey Level Group" }).Description

#Write-Host $otherAssessmentDescription

$costDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Cost Optimization" -and $_.Category -eq "Survey Level Group" }).Description
$operationsDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Operational Excellence" -and $_.Category -eq "Survey Level Group" }).Description
$performanceDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Performance Efficiency" -and $_.Category -eq "Survey Level Group" }).Description
$reliabilityDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Reliability" -and $_.Category -eq "Survey Level Group" }).Description
$securityDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Security" -and $_.Category -eq "Survey Level Group" }).Description
function Get-PillarInfo($pillar) {
    if ($pillar.Contains("Cost Optimization")) {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $costScore; "Description" = $costDescription; "ScoreDescription" = $OverallScoreDescription }
    }
    if ($pillar.Contains("Reliability")) {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $reliabilityScore; "Description" = $reliabilityDescription; "ScoreDescription" = $ReliabilityScoreDescription }
    }
    if ($pillar.Contains("Operational Excellence")) {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $operationsScore; "Description" = $operationsDescription; "ScoreDescription" = $OperationsScoreDescription }
    }
    if ($pillar.Contains("Performance Efficiency")) {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $performanceScore; "Description" = $performanceDescription; "ScoreDescription" = $PerformanceScoreDescription }
    }
    if ($pillar.Contains("Security")) {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $securityScore; "Description" = $securityDescription; "ScoreDescription" = $SecurityScoreDescription }
    }
}
#endregion

#region Instantiate PowerPoint variables
#write-debug "Run PowerPoint"
$application = New-Object -ComObject powerpoint.application
$application.visible = -1 # [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $application.Presentations.open($templatePresentation)

if ($assessmentTypeCheck.contains("Well-Architected")) {
    $titleSlide = $presentation.Slides[8]
    $summarySlide = $presentation.Slides[9]
    $detailSlide = $presentation.Slides[10]
    $endSlide = $presentation.Slides[11]
}
else {
    $titleSlide = $presentation.Slides[7]
    $summarySlide = $presentation.Slides[8]
    $detailSlide = $presentation.Slides[9]
    $endSlide = $presentation.Slides[10]
}

#endregion



if ($assessmentTypeCheck.contains("Well-Architected")) {

    foreach ($pillar in $filteredpillars) {
        $pillarData = $data | Where-Object { $_.Category -eq $pillar }
        # for debug
        #Write-host -Debug $data

        $pillarInfo = Get-PillarInfo -pillar $pillar
        # Write-host -Debug "PILLAR INFO: $pillarInfo"
        # Edit title & date on slide 1
        $slideTitle = $title.Replace("[pillar]", $pillar) #,$pillar.substring(0,1).toupper()+$pillar.substring(1).tolower()) #lowercase only here?
        $newTitleSlide = $titleSlide.Duplicate()
        $newTitleSlide.MoveTo($presentation.Slides.Count)
        $newTitleSlide.Shapes[3].TextFrame.TextRange.Text = $slideTitle
        $newTitleSlide.Shapes[4].TextFrame.TextRange.Text = $newTitleSlide.Shapes[4].TextFrame.TextRange.Text.Replace("[Report_Date]", $localReportDate)

        # Edit Executive Summary Slide

        #Add logic to get overall score
        $newSummarySlide = $summarySlide.Duplicate()
        $newSummarySlide.MoveTo($presentation.Slides.Count)
        $newSummarySlide.Shapes[3].TextFrame.TextRange.Text = $pillarInfo.Score
        $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $pillarInfo.Description
        [Double]$summBarScore = [int]$pillarInfo.Score * 2.47 + 56
        $newSummarySlide.Shapes[11].Left = $summBarScore

        $CategoriesList = New-Object System.Collections.ArrayList
        $categories = ($pillarData | Sort-Object -Property "Weight" -Descending).ReportingCategory | Select-Object -Unique
        foreach ($category in $categories) {
            $categoryWeight = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
            $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
            $categoryWeightiestCount = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight -ge $MinimumReportLevel | Measure-Object
            $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count }) | Out-Null
        }

        $CategoriesList = $CategoriesList | Sort-Object -Property CategoryScore -Descending
        write-host -Debug $CategoriesList
        $counter = 13 #Shape count for the slide to start adding scores
        $categoryCounter = 0
        $areaIconX = 378.1129
        $areaIconY = @(176.4359, 217.6319, 258.3682, 299.1754, 339.8692, 382.6667, 423.9795, 461.0491)
        foreach ($category in $CategoriesList) {
            if ($category.Category -ne "Uncategorized") {
                try {
                    #$newSummarySlide.Shapes[8] #Domain 1 Icon
                    $newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryWeightiestCount.ToString("#")
                    $newSummarySlide.Shapes[$counter + 1].TextFrame.TextRange.Text = $category.Category
                    $counter = $counter + 3
                    switch ($category.CategoryScore) {
                        { $_ -lt 33 } { 
                            $categoryShape = $newSummarySlide.Shapes[37]
                        }
                        { $_ -gt 33 -and $_ -lt 67 } { 
                            $categoryShape = $newSummarySlide.Shapes[38] 
                        }
                        { $_ -gt 67 } { 
                            $categoryShape = $newSummarySlide.Shapes[39] 
                        }
                        Default { 
                            $categoryShape = $newSummarySlide.Shapes[38] 
                        }
                    }
                    $categoryShape.Duplicate() | Out-Null
                    $newShape = $newSummarySlide.Shapes.Count
                    $newSummarySlide.Shapes[$newShape].Left = $areaIconX
                    $newSummarySlide.Shapes[$newShape].top = $areaIcony[$categoryCounter] 
                    $categoryCounter = $categoryCounter + 1
                }
                catch {}
            }
        }

        #Remove the boilerplate placeholder text if categories < 8
        if ($categories.Count -lt 8) {
            for ($k = $newSummarySlide.Shapes.count; $k -gt $counter - 1; $k--) {
                try {
                    $newSummarySlide.Shapes[$k].Delete()
                    <#$newSummarySlide.Shapes[$k].Delete()
                    $newSummarySlide.Shapes[$k+1].Delete()#>
                }
                catch {}
            }
        }

        # Edit new category summary slide

        foreach ($category in $CategoriesList.Category) {
            #Write-Debug "Processing Category: $category"
            $categoryData = $pillarData | Where-Object { $_.ReportingCategory -eq $category -and $_.Category -eq $pillar }
            $categoryDataCount = ($categoryData | measure).Count
            #write-Debug "  Category Data Count: $categoryDataCount"
            $categoryWeight = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
            $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
            $categoryDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq $pillar -and $categoryData.ReportingCategory.Contains($_.Category) }).Description
            $y = $categoryDataCount
            $x = $ShowTop
            if ($categoryDataCount -lt $x) {
                $x = $categoryDataCount
            }

            $newDetailSlide = $detailSlide.Duplicate()
            $newDetailSlide.MoveTo($presentation.Slides.Count)

            $newDetailSlide.Shapes[1].TextFrame.TextRange.Text = $category
            $newDetailSlide.Shapes[3].TextFrame.TextRange.Text = $categoryScore.ToString("#")
            [Double]$detailBarScore = $categoryScore * 2.48 + 38
            $newDetailSlide.Shapes[12].Left = $detailBarScore
            $newDetailSlide.Shapes[4].TextFrame.TextRange.Text = $categoryDescription
            $newDetailSlide.Shapes[7].TextFrame.TextRange.Text = "Top $x out of $y recommendations:"
            $newDetailSlide.Shapes[8].TextFrame.TextRange.Text = ($categoryData | Sort-Object -Property "Link-Text" -Unique | Sort-Object -Property Weight -Descending | Select-Object -First $x).'Link-Text' -join "`r`n`r`n"
            $sentenceCount = $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences().count

            for ($k = 1; $k -le $sentenceCount; $k++) {
                if ($newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).Text) {
                    try {
                        $recommendationObject = $categoryData | Where-Object { $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).Text.Contains($_.'Link-Text') }
                        $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).ActionSettings(1).HyperLink.Address = $recommendationObject.Link
                    }
                    catch {}
                }
            }    
        }

    }
} 
else { #Assessment type is NOT Well-Architected
    $slideTitle = $title.Replace("[CA_Security_Review]", "Cloud Adoption Security Review")
    $newTitleSlide = $titleSlide.Duplicate()
    $newTitleSlide.MoveTo($presentation.Slides.Count)
    $newTitleSlide.Shapes[3].TextFrame.TextRange.Text = $slideTitle
    $newTitleSlide.Shapes[4].TextFrame.TextRange.Text = $newTitleSlide.Shapes[4].TextFrame.TextRange.Text.Replace("[Report_Date]", $localReportDate)

    # Edit Executive Summary Slide
    if (![string]::IsNullOrEmpty($overallScore)) {
        $ScoreText = "$($overallScore)"
    }


    #Add logic to get overall score
    $newSummarySlide = $summarySlide.Duplicate()
    $newSummarySlide.MoveTo($presentation.Slides.Count)
    $newSummarySlide.Shapes[3].TextFrame.TextRange.Text = $ScoreText
    $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $otherAssessmentDescription
    [Double]$summBarScore = [int]$ScoreText * 2.47 + 56
    $newSummarySlide.Shapes[11].Left = $summBarScore


    $CategoriesList = New-Object System.Collections.ArrayList
    #Updated to use ReportingCategory vs Category due to Category column for CASR containing multiple instances of varying interests vs WASA(ie. "Security")
    $categories = $data.ReportingCategory | Sort-Object -Property "Weight" -Descending | Select-Object -Unique
    foreach ($category in $categories) {
        $categoryWeight = ($data | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
        $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
        $categoryWeightiestCount = ($data | Where-Object { $_.ReportingCategory -eq $category }).Weight -ge $MinimumReportLevel | Measure-Object
        $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count }) | Out-Null
    }

    $CategoriesList = $CategoriesList | Sort-Object -Property CategoryScore -Descending

    $counter = 13 #Shape count for the slide to start adding scores
    $categoryCounter = 0
    $areaIconX = 378.1129
    $areaIconY = @(176.4359, 217.6319, 258.3682, 299.1754, 339.8692, 382.6667, 423.9795, 461.0491)

    foreach ($category in $CategoriesList) {
        if ($category.Category -ne "Uncategorized") {
            try {
                #$newSummarySlide.Shapes[8] #Domain 1 Icon
                #$newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryScore.ToString("#")
                $newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryWeightiestCount.ToString("#")
                $newSummarySlide.Shapes[$counter + 1].TextFrame.TextRange.Text = $category.Category
                $counter = $counter + 3 # no graphic anymore
                switch ($category.CategoryScore) {
                    { $_ -lt 33 } { 
                        $categoryShape = $newSummarySlide.Shapes[37]
                    }
                    { $_ -gt 33 -and $_ -lt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[38] 
                    }
                    { $_ -gt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[39] 
                    }
                    Default { 
                        $categoryShape = $newSummarySlide.Shapes[38] 
                    }
                }
                $categoryShape.Duplicate() | Out-Null
                $newShape = $newSummarySlide.Shapes.Count
                $newSummarySlide.Shapes[$newShape].Left = $areaIconX
                $newSummarySlide.Shapes[$newShape].top = $areaIcony[$categoryCounter] 
                $categoryCounter = $categoryCounter + 1
            }
            catch {}
        }
    }

    #Remove the boilerplate placeholder text if categories < 8
    if ($categories.Count -lt 8) {
        for ($k = $newSummarySlide.Shapes.count; $k -gt $counter - 1; $k--) {
            try {
                $newSummarySlide.Shapes[$k].Delete()
                <#$newSummarySlide.Shapes[$k].Delete()
            $newSummarySlide.Shapes[$k+1].Delete()#>
            }
            catch {}
        }
    }

    # Edit new category summary slide

    foreach ($category in $CategoriesList.Category) {

        $categoryData = $data | Where-Object { $_.ReportingCategory -eq $category }# -and $_.Category -eq $casr}
        $categoryDataCount = ($categoryData | Measure-Object).Count
        $categoryWeight = ($data | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
        $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
        $categoryDescription = ($descriptionsFile | Where-Object { $categoryData.ReportingCategory.Contains($_.Category) }).Description
        $y = $categoryDataCount
        $x = $ShowTop
        if ($categoryDataCount -lt $x) {
            $x = $categoryDataCount
        }

        $newDetailSlide = $detailSlide.Duplicate()
        $newDetailSlide.MoveTo($presentation.Slides.Count)

        $newDetailSlide.Shapes[1].TextFrame.TextRange.Text = $category
        $newDetailSlide.Shapes[3].TextFrame.TextRange.Text = $categoryScore.ToString("#")
        [Double]$detailBarScore = $categoryScore * 2.48 + 38
        $newDetailSlide.Shapes[12].Left = $detailBarScore
        $newDetailSlide.Shapes[4].TextFrame.TextRange.Text = $categoryDescription
        $newDetailSlide.Shapes[7].TextFrame.TextRange.Text = "Top $x out of $y recommendations:"
        $newDetailSlide.Shapes[8].TextFrame.TextRange.Text = ($categoryData | Sort-Object -Property "Link-Text" -Unique | Sort-Object -Property Weight -Descending | Select-Object -First $x).'Link-Text' -join "`r`n`r`n"
        $sentenceCount = $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences().count

        for ($k = 1; $k -le $sentenceCount; $k++) {
            if ($newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).Text) {
                try {
                    $recommendationObject = $categoryData | Where-Object { $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).Text.Contains($_.'Link-Text') }
                    $newDetailSlide.Shapes[8].TextFrame.TextRange.Sentences($k).ActionSettings(1).HyperLink.Address = $recommendationObject.Link
                }
                catch {}
            }
        }     
    }
}

$newEndSlide = $endSlide.Duplicate()
$newEndSlide.MoveTo($presentation.Slides.Count)

$titleSlide.Delete()
$summarySlide.Delete()
$detailSlide.Delete()
$endSlide.Delete()

if ($assessmentTypeCheck.contains("Well-Architected")) {
    $presentation.SavecopyAs("$workingDirectory\WAF-Review-$($reportDate).pptx")
}
else {
    $presentation.SavecopyAs("$workingDirectory\CASR-$($reportDate).pptx")
}

$presentation.Close()


$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()
