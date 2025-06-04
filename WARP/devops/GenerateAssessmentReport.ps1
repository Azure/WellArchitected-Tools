#requires -Version 7
<#
.SYNOPSIS
    Takes output from the Well-Architected Review Assessment website and produces a PowerPoint presentation incorporating the findings.
    Also support the Cloud Adoption Security Assessment and the DevOps Capability Assessment.
    https://learn.microsoft.com/en-us/assessments/azure-architecture-review/
    
.DESCRIPTION
    The Well-Architected Review site provides a self-Assessment tool. This might be used for self-assessment, or run as part of an Assessment performed by Microsoft to help recommend improvements.

    From the directory in which you've downloaded the scripts, templates and csv file from the PnP survey (mycontent.csv)
    For a WAF Assessment report:

        .\GenerateAssessmentReport.ps1 -ContentFile .\mycontent.csv

    For a CASA report:

        .\GenerateAssessmentReport.ps1 -ContentFile .\mycontent.csv -CloudAdoption

    For a DevOps capability report:

        .\GenerateAssessmentReport.ps1 -ContentFile .\mycontent.csv -DevOpsCapability

    Ensure the powerpoint template file and the Category Descriptions file exist in the paths shown below before attempting to run this script
    Once the script is run, close the powershell window and a timestamped PowerPoint report and a subset csv file will be created on the working directory
    Use these reports to represent and edit your findings for the WAF Engagement

    Known issues 
    - If the hyperlinks are not being published accurately, ensure that the csv file doesnt have any multi-sentence recommendations under Link-Text field


.PARAMETER ContentFile
    Exported CSV file from the Well-Architected Review file. Supports relative paths.

.PARAMETER CloudAdoption
    If set, indicates the Cloud Adoption Security Assessment format should be used. If not set, Well-Architected is assumed.

.PARAMETER DevOpsCapability
    If set, indicates the DevOps Capability Review format should be used. If not set, Well-Architected is assumed.

.PARAMETER MinimumReportLevel
    The level above which a finding is considered high severity, By convention, scores up to 32 are low, 65 medium, and 66+ high.

.PARAMETER ShowTop
    How many recommendations to try to fit on a slide. 8 is default.

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
    .\generateAssessmentReport.ps1 -ContentFile .\Cloud_Adoption_Security_Assessment_Sample.csv -ShowTop 9
    
    Generates a PPTX report from a CASA CSV
    
    Tries to include the top 9 results of any category (which probably won't fit by default, so plan to reformat things)
        
.EXAMPLE
    .\generateAssessmentReport.ps1 -ContentFile .\Cloud_Adoption_Security_Assessment_Sample.csv -CloudAdoption    
    
    If the title doesn't identify the report type correctly, you can force the decision with the relevant switch. (for example: -CloudAdoption, -DevOpsCapability)

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

    # Show Top N Recommendations Per Slide (default 5)
    [Parameter()][int]
    $ShowTop = 5 ,

    [Parameter()]
    [switch] $CloudAdoption,

    [Parameter()]
    [switch] $DevOpsCapability

)


#region Functions

$assessmentFile = ""

function OpenAssessmentFile
{

    if ($null -eq $ContentFile -or !(Test-Path $ContentFile)) {
        $inputFile = Get-FileName $workingDirectory
    } else {
        $inputFile = $ContentFile
    }

    # validate our file is OK
    try {
        $content = Get-Content  $inputFile
    }
    catch {
        Write-Error -Message "Unable to open selected Content file."
        exit
    }

    $global:assessmentFile = $inputFile
    return $content 

}


function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.Title = "Select review file export"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

function FindIndexBeginningWith($stringset, $searchterm) {
    $i = 0
    foreach ($line in $stringset) {
        if ($line.StartsWith($searchterm)) {
            return $i
        }
        $i++
    }
    return false
}

function LoadDescriptionFile {
    if ($WellArchitected) {
        try {
            $descriptionsFile = Import-Csv "$workingDirectory\WAF Category Descriptions.csv"
        }
        catch {
            Write-Error -Message "Unable to open $($workingDirectory)\WAF Category Descriptions.csv"
            exit
        }
    }
    elseif ($DevOpsCapability) {
        try {
            $descriptionsFile = Import-Csv "$workingDirectory\DevOps Category Descriptions.csv"
        }
        catch {
            Write-Error -Message "Unable to open $($workingDirectory)\DevOps Category Descriptions.csv"
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
    return $descriptionsFile

}

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

function GetMappedReportingCategory{
    param (
        $reportingCategrory,
        $currentPillar
    )

    $newReportingCategory = ($descriptionsFile | Where-Object { $_.Pillar -eq $currentPillar -and $_.Category.StartsWith($reportingCategrory) }).Caption
    if (-not $newReportingCategory) {
        $newReportingCategory = $reportingCategrory # Fallback to existing ReportingCategory if no mapping found
    }

    return $newReportingCategory
}



Function WellArchitectedAssessment
{
    foreach ($pillar in $filteredpillars) 
    {
        $pillarData = $data | Where-Object { $_.Category -eq $pillar }
        # for debug: Write-host -Debug $data

        $pillarInfo = Get-PillarInfo -pillar $pillar
        # Write-host -Debug "PILLAR INFO: $pillarInfo"
        # Populates Title Slide

        $slideTitle = $title.Replace("[pillar]", $pillar) #,$pillar.substring(0,1).toupper()+$pillar.substring(1).tolower()) #lowercase only here?
        $newTitleSlide = $titleSlide.Duplicate()
        $newTitleSlide.MoveTo($presentation.Slides.Count)
        $newTitleSlide.Shapes[3].TextFrame.TextRange.Text = $slideTitle
        $newTitleSlide.Shapes[4].TextFrame.TextRange.Text = $newTitleSlide.Shapes[4].TextFrame.TextRange.Text.Replace("[Report_Date]", $localReportDate)


        # Populates Executive Summary Slide(s)

        # prepare category list and identify "high importance" recommendations

        $CategoriesList = New-Object System.Collections.ArrayList
        $categories = ($pillarData | Sort-Object -Property "Weight" -Descending).ReportingCategory | Select-Object -Unique
        foreach ($category in $categories) {
            $categoryWeight = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
            $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
            $categoryWeightiestCount = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight -ge $MinimumReportLevel | Measure-Object
            $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count }) | Out-Null
        }

        # display categories alphabetically - so that the WAF 2.0 code numbers are in order
        $CategoriesList = $CategoriesList | Sort-Object -Property Category

        $newSummarySlide = $summarySlide.Duplicate()
        $newSummarySlide.MoveTo($presentation.Slides.Count)
        $newSummarySlide.Shapes[3].TextFrame.TextRange.Text = $pillarInfo.Score
        $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $pillarInfo.Description
        [Double]$summBarScore = [int]$pillarInfo.Score * 2.47 + 56
        $newSummarySlide.Shapes[11].Left = $summBarScore

        $counter = 13 #Shape index for the slide to start adding scores (it's 13 because the first 12 shapes are the boilerplate text: 12 is gauge icon, 13 is score, 14 is category, etc.)
        $categoryCounter = 0
        $pageCounter = 1
        $gaugeIconX = 437.76 # X coordinate for the gauge icon in points (1 point = 1/72 inch)
        $gaugeIconY = @(147.6, 176.4, 204.48, 232.56, 261.36, 289.44, 317.52, 346.32, 375.12, 403.2, 432.0, 460.08 ) # Y coordinates for the remaining 12 gauge icons in points (vertically aligned)
        
        # Filter out any empty / non-existing categories including "Uncategorized" (aka Advisor)
        $FilteredCategoriesList = ($CategoriesList | Where-Object { $_.Category -ne "" -and $_.Category -ne "Uncategorized" })
        $CategoriesList = $FilteredCategoriesList 
        
        foreach ($category in $CategoriesList) {
            if($categoryCounter -ge (12 * $pageCounter)) {
                # add another page if there are more categories than can fit
                $newSummarySlide = $summarySlide.Duplicate()
                $newSummarySlide.MoveTo($presentation.Slides.Count)
                $newSummarySlide.Shapes[3].TextFrame.TextRange.Text = $pillarInfo.Score
                $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $pillarInfo.Description
                [Double]$summBarScore = [int]$pillarInfo.Score * 2.47 + 56
                $newSummarySlide.Shapes[11].Left = $summBarScore

                $counter = 13 #Shape count for the slide to start adding scores
                $categoryCounter = 0
                $pageCounter = $pageCounter + 1
                $gaugeIconX = 437.76 # X coordinate for the gauge icon in points (1 point = 1/72 inch)
                $gaugeIconY = @(147.6, 176.4, 204.48, 232.56, 261.36, 289.44, 317.52, 346.32, 375.12, 403.2, 432.0, 460.08 ) # Y coordinates for the remaining 12 gauge icons in points (vertically aligned)
            }

            try {
                $newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryWeightiestCount.ToString("0")
                    
                # Replacing the domain area (aka category) with the caption (aka new category) from the description file (if any)
                $newSummarySlide.Shapes[$counter + 1].TextFrame.TextRange.Text = GetmappedReportingCategory -reportingCategrory $category.Category -currentPillar $pillar

                $counter = $counter + 3 # select the next score textbox shape on the slide (1. Gauge, 2. Score, 3. Category)
                    
                # Determining the color based on CategoryScore
                switch ($category.CategoryScore) {
                    { $_ -lt 33 } { 
                        $categoryShape = $newSummarySlide.Shapes[49] #green
                        break
                    }
                    { $_ -gt 33 -and $_ -lt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[50] #yellow
                        break
                    }
                    { $_ -gt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[51] #red
                        break
                    }
                    Default { 
                        $categoryShape = $newSummarySlide.Shapes[50] #yellow
                    }
                }
                $categoryShape.Duplicate() | Out-Null
                $newShape = $newSummarySlide.Shapes.Count
                $newSummarySlide.Shapes[$newShape].Left = $gaugeIconX
                $newSummarySlide.Shapes[$newShape].top = $gaugeIconY[$categoryCounter] 
                $newSummarySlide.Shapes[$newShape].Name = "NewGauges"

                $categoryCounter = $categoryCounter + 1
            }
            catch {
                Write-Error "Error: $($Error[0].Exception.ToString())" 
                Write-Error "Error: $($Error[0].Exception.InnerException.ToString())"
                continue
            }
        }

        # Populates Pillar Slides

        foreach ($category in $CategoriesList.Category) {
            #Write-Debug "Processing Category: $category"
            $categoryData = $pillarData | Where-Object { $_.ReportingCategory -eq $category -and $_.Category -eq $pillar }
            $categoryDataCount = ($categoryData | measure).Count
            #write-Debug "  Category Data Count: $categoryDataCount"
            $categoryWeight = ($pillarData | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
            $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
            $categoryDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq $pillar -and $_.Category.StartsWith($category) }).Description
            # Replacing the domain area (aka category) with the caption (aka new category) from the description file (if any)
            $categoryTitle = GetmappedReportingCategory -reportingCategrory $category -currentPillar $pillar

            $y = $categoryDataCount
            $x = $ShowTop
            if ($categoryDataCount -lt $x) {
                $x = $categoryDataCount
            }

            $newDetailSlide = $detailSlide.Duplicate()
            $newDetailSlide.MoveTo($presentation.Slides.Count)

            $newDetailSlide.Shapes[1].TextFrame.TextRange.Text = $categoryTitle
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

        #Remove boilerplate shapes. 
        if ($categories.Count -lt (12 * $pageCounter)) { #12 is the number of categories shapes on the summary slide 
            for ($k = $newSummarySlide.Shapes.count; $k -gt $counter - 1; $k--) {
                if ($null -ne $newSummarySlide.Shapes[$k] -and $newSummarySlide.Shapes[$k].Name -ne "NewGauges") { # Fix: Don't delete the newly added colored gauge shapes 
                    try {
                        $newSummarySlide.Shapes[$k].Delete()
                    }
                    catch {}
                }
            }
        }
    }
}

Function CloudAdoptionAssessment
{
    $slideTitle = $title.Replace("[CAF_Security_Assessment]", "Cloud Adoption Security Assessment")
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
    $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $cloudAdoptionDescription
    [Double]$summBarScore = [int]$ScoreText * 2.47 + 56
    $newSummarySlide.Shapes[11].Left = $summBarScore


    $CategoriesList = New-Object System.Collections.ArrayList
    #Updated to use ReportingCategory vs Category due to Category column for CASA containing multiple instances of varying interests vs WASA(ie. "Security")
    $categories = $data.ReportingCategory | Sort-Object -Property "Weight" -Descending | Select-Object -Unique
    
        
    # Remove non existing (aka empty) categories. CASA has only 6 categories (no Advisor/uncategorized category)
    $FilteredCategoriesList = [System.Collections.ArrayList]($categories | Where-Object { $_ -ne "" })
    $categories = $FilteredCategoriesList
    
    foreach ($category in $categories) {
        $categoryWeight = ($data | Where-Object { $_.ReportingCategory -eq $category }).Weight | Measure-Object -Sum
        $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
        $categoryWeightiestCount = ($data | Where-Object { $_.ReportingCategory -eq $category }).Weight -ge $MinimumReportLevel | Measure-Object
        $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count }) | Out-Null
    }

    $CategoriesList = $CategoriesList | Sort-Object -Property CategoryScore -Descending

    $counter = 13 #Shape count for the slide to start adding scores
    $categoryCounter = 0
    $gaugeIconX = 437.76 # X coordinate for the gauge icon in points (1 point = 1/72 inch)
    $gaugeIconY = @(147.6, 176.4, 204.48, 232.56, 261.36, 289.44, 317.52, 346.32, 375.12, 403.2, 432.0, 460.08 ) # Y coordinates for the remaining 12 gauge icons in points (vertically aligned)

    foreach ($category in $CategoriesList) {
        if ($category.Category -ne "Uncategorized") {
            try {
                #$newSummarySlide.Shapes[8] #Domain 1 Icon
                #$newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryScore.ToString("#")
                $newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryWeightiestCount.ToString("#")
                $newSummarySlide.Shapes[$counter + 1].TextFrame.TextRange.Text = $category.Category
                $counter = $counter + 3 # no graphic anymore
                # Determining the color based on CategoryScore
                switch ($category.CategoryScore) {
                    { $_ -lt 33 } { 
                        $categoryShape = $newSummarySlide.Shapes[49] #green
                        break
                    }
                    { $_ -gt 33 -and $_ -lt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[50] #yellow
                        break
                    }
                    { $_ -gt 67 } { 
                        $categoryShape = $newSummarySlide.Shapes[51] #red
                        break
                    }
                    Default { 
                        $categoryShape = $newSummarySlide.Shapes[50] #yellow
                    }
                }
                $categoryShape.Duplicate() | Out-Null
                $newShape = $newSummarySlide.Shapes.Count
                $newSummarySlide.Shapes[$newShape].Left = $gaugeIconX
                $newSummarySlide.Shapes[$newShape].top = $gaugeIconY[$categoryCounter] 
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
                $newSummarySlide.Shapes[$k+1].Delete()
            }
            catch {}
        }
    }

    # Edit new category summary slide

    foreach ($category in $CategoriesList.Category) {

        $categoryData = $data | Where-Object { $_.ReportingCategory -eq $category }
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

Function DevOpsCapabilityAssessment {
    $slideTitle = $title.Replace("[CA_Security_Review]", "DevOps Capability Review")
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
    $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $devOpsDescription
    [Double]$summBarScore = [int]$ScoreText * 2.47 + 56
    $newSummarySlide.Shapes[11].Left = $summBarScore


    $CategoriesList = New-Object System.Collections.ArrayList
    $categories = $data.Category | Sort-Object -Property "Weight" -Descending | Select-Object -Unique
    
        
    # Remove non existing (aka empty) categories. CASA has only 6 categories (no Advisor/uncategorized category)
    $FilteredCategoriesList = [System.Collections.ArrayList]($categories | Where-Object { $_ -ne "" })
    $categories = $FilteredCategoriesList
    
    foreach ($category in $categories) {
        $categoryWeight = ($data | Where-Object { $_.Category -eq $category }).Weight | Measure-Object -Sum
        $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
        $categoryWeightiestCount = ($data | Where-Object { $_.Category -eq $category }).Weight -ge $MinimumReportLevel | Measure-Object
        $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count }) | Out-Null
    }

    $CategoriesList = $CategoriesList | Sort-Object -Property CategoryScore -Descending

    $counter = 13 #Shape count for the slide to start adding scores
    $categoryCounter = 0
    $gaugeIconX = 378.1129
    $gaugeIconY = @(176.4359, 217.6319, 258.3682, 299.1754, 339.8692, 382.6667, 423.9795, 461.0491)

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
                $newSummarySlide.Shapes[$newShape].Left = $gaugeIconX
                $newSummarySlide.Shapes[$newShape].top = $gaugeIconY[$categoryCounter] 
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
                $newSummarySlide.Shapes[$k+1].Delete()
            }
            catch {}
        }
    }

    # Edit new category summary slide

    foreach ($category in $CategoriesList.Category) {

        $categoryData = $data | Where-Object { $_.Category -eq $category }
        $categoryDataCount = ($categoryData | Measure-Object).Count
        $categoryWeight = ($data | Where-Object { $_.Category -eq $category }).Weight | Measure-Object -Sum
        $categoryScore = $categoryWeight.Sum / $categoryWeight.Count
        $categoryDescription = ($descriptionsFile | Where-Object { $categoryData.Category.Contains($_.Category) }).Description
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


Function CleanUp
{

    try {
        $newEndSlide = $endSlide.Duplicate()
        $newEndSlide.MoveTo($presentation.Slides.Count)
        $titleSlide.Delete()
        $summarySlide.Delete()
        $detailSlide.Delete()
        $endSlide.Delete()        
    }
    catch {
    }

    if ($WellArchitected)
    {
        $presentation.SavecopyAs("$workingDirectory\WAF-Review-$($reportDate).pptx")
    }
    elseif($DevOpsCapability){
        $presentation.SavecopyAs("$workingDirectory\DevOps-$($reportDate).pptx")
    } 
    else {
        $presentation.SavecopyAs("$workingDirectory\CASA-$($reportDate).pptx")
    }

    $presentation.Close()
    $application.quit()
    $application = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
    

}

#endregion


#region Main


$workingDirectory = (Get-Location).Path #Get the working directory from the script
$content  = OpenAssessmentFile
$assessmentTypeCheck = ""
$assessmentTypeCheck = ($content | Select-Object -First 1)

$reportDate = Get-Date -Format "yyyy-MM-dd-HHmm"
$localReportDate = Get-Date -Format g
$overallScore = ""
$costScore = ""
$operationsScore = ""
$performanceScore = ""
$reliabilityScore = ""
$securityScore = ""
$overallScoreDescription = ""

$filteredPillars = @()

$WellArchitected  = $false
$CloudAdoption    = $false
$DevOpsCapability = $false

# Respect user switches first
if ($PSBoundParameters.ContainsKey('CloudAdoption')) {
    write-host "-CloudAdoption switch detected"
    $CloudAdoption    = $true
}
elseif ($PSBoundParameters.ContainsKey('DevOpsCapability')) {
    write-host "-DevOpsCapability switch detected"
    $DevOpsCapability = $true
}
else {
    # Only if no switch: auto-detect from CSV title
    if ($assessmentTypeCheck.Contains('Cloud Adoption')) {
        write-host "Auto detected Cloud Adoption Security Assessment CSV file"
        $CloudAdoption   = $true
    }
    elseif ($assessmentTypeCheck.Contains('DevOps Capability')) {
        write-host "Auto detected DevOps Capability Review CSV file"
        $DevOpsCapability = $true
    }
    else {
        $WellArchitected = $true
    }
}

if ($WellArchitected) {
    for ($i = 3; $i -le 8; $i++) {

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

if ($WellArchitected) {
    Write-host "Producing Well Architected report from  $global:assessmentFile"
    $templatePresentation = "$workingDirectory\PnP_PowerPointReport_Template.pptx"
    $title = "Well-Architected [pillar] Assessment" # Don't edit this - it's used when multiple Pillars are included.
    try {
        $tableStart = FindIndexBeginningWith $content "Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context"
    }
    catch{
        Write-host "That appears not to be a content file. Please use only content from the Well-Architected Assessment site."
    }
    try{
        
        $EndStringIdentifier = $content | Where-Object { $_.Contains("--,,") } | Select-Object -Unique -First 1
        
        $tableEnd = $content.IndexOf($EndStringIdentifier) - 1
        
        $csv = $content[$tableStart..$tableEnd] | Out-File  "$workingDirectory\$reportDate.csv"
        $importdata = Import-Csv -Path "$workingDirectory\$reportDate.csv"

        # Clean the uncategorized data
        
            foreach ($lineData in $importdata) {
                
                if (!$lineData.ReportingCategory) {
                    $lineData.ReportingCategory = "Uncategorized"
                }
            }

        $data = $importdata | where {$_.Category -in $filteredPillars}
        $data | Export-Csv -UseQuotes AsNeeded "$workingDirectory\$reportDate.csv" 
        $data | % { $_.Weight = [int]$_.Weight }
        $pillars = $data.Category | Select-Object -Unique
    }
    catch {
        Write-Host "Unable to parse the content file."
        Write-Host "Please ensure all input files are in the correct format and aren't open in Excel or another editor which locks the file."
        Write-Host "--"
        Write-Host $_
        exit
    }
} else {
    if($CloudAdoption) {
        Write-host "Producing Cloud Adoption Security Assessment report from $global:assessmentFile"
        $templatePresentation = "$workingDirectory\PnP_PowerPointReport_Template - CAF-Secure.pptx"
        $title = "Cloud Adoption Security Assessment"
    } elseif($DevOpsCapability) {
        Write-host "Producing DevOps Capability Review report from $global:assessmentFile"
        $templatePresentation = "$workingDirectory\PnP_PowerPointReport_Template - DevOps.pptx"
        $title = "DevOps Capability Review"
    }

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


$descriptionsFile = LoadDescriptionFile

$cloudAdoptionDescription = ($descriptionsFile | Where-Object { $_.Category -eq "Survey Level Group" }).Description
$devOpsDescription = ($descriptionsFile | Where-Object { $_.Category -eq "Survey Level Group" }).Description
$costDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Cost Optimization" -and $_.Category -eq "Survey Level Group" }).Description
$operationsDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Operational Excellence" -and $_.Category -eq "Survey Level Group" }).Description
$performanceDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Performance Efficiency" -and $_.Category -eq "Survey Level Group" }).Description
$reliabilityDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Reliability" -and $_.Category -eq "Survey Level Group" }).Description
$securityDescription = ($descriptionsFile | Where-Object { $_.Pillar -eq "Security" -and $_.Category -eq "Survey Level Group" }).Description


#region Instantiate PowerPoint variables

$application = New-Object -ComObject powerpoint.application
$application.visible = -1 # [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $application.Presentations.open($templatePresentation)

if ($WellArchitected) {
    $titleSlide = $presentation.Slides[9]
    $summarySlide = $presentation.Slides[10]
    $detailSlide = $presentation.Slides[11]
    $endSlide = $presentation.Slides[12]
}
else {
    $titleSlide = $presentation.Slides[3]
    $summarySlide = $presentation.Slides[4]
    $detailSlide = $presentation.Slides[5]
    $endSlide = $presentation.Slides[6]
}

#endregion


if ($WellArchitected) 
{
    WellArchitectedAssessment
}
elseif ($DevOpsCapability) 
{
    DevOpsCapabilityAssessment
}
else 
{
    CloudAdoptionAssessment
}

CleanUp

#endregion