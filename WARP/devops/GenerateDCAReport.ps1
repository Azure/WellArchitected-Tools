[CmdletBinding()]
param (
    # Indicates CSV file for input
        [Parameter()][string]
    $ContentFile = "DevOps_Capability_Assessment_NBSimul.csv" ,
    #"C:\**\DevOps_Capability_Assessment*.csv"  ,

    # Set working directory for output files
        [Parameter()][string]
    $WorkingDir = "C:\**"  ,

    # Presentation filename to build from
        [Parameter()][string]
    $templatePresentationFile = "C:\*\PnP_PowerPointReport_Template.pptx"  ,

    # Descriptions File
        [Parameter()][string]
    $DescriptionsCSVFile = "C:\U*\DOCACategoryDescriptions.csv"  ,

    # Minimum level for inclusion in summary (defaults to High)
        [Parameter()][int]
    $MinimumReportLevel = 65   ,

    # Show Top N Recommendations Per Slide (default 8)
    [Parameter()][int]
    $ShowTop = 8   

)
<# Instructions to use this script

1. Set the workingDirectory value in the script to a folder path that includes the scripts, templates and the downloaded csv file from PnP
2. Set the right csv file name on $content value and point it to the downloaded csv file path
3. Ensure the powerpoint template file and the Category Descriptions file exist in the paths shown below before attempting to run this script
4. Once the script is run, close the powershell window and a timestamped PowerPoint report and a subset csv file will be created on the working directory
5. Use these reports to represent and edit your findings for the DOCA Engagement
6. Known issues 
    a. Practice scores may not reflect accurately if the ordering in the csv is jumbled. Please adjust lines 41-53 in case the score representations for the practices are not accurate
    b. If the hyperlinks are not being published accurately, ensure that the csv file doesnt have any multi-sentence recommendations under Link-Text field

#>

$WorkingDirectory = Resolve-Path $WorkingDir
$content = Get-Content $ContentFile
$descriptionsFile = Import-Csv $DescriptionsCSVFile
$templatePresentation = Resolve-Path $templatePresentationFile

$title = "DevOps Capability Assessment for [pillar]" # Don't edit this - it's used when multiple Pillars are included.
$reportDate = Get-Date -Format "yyyy-MM-dd-HHmm"
$localReportDate = Get-Date -Format g
#$tableStart = $content.IndexOf("Title,Description,Link-Text,Link,Priority,Category,Subcategory,Weight")
$tableStart = $content.IndexOf("Category,Link-Text,Link,Priority,ReportingCategory,ReportingSubcategory,Weight,Context")
$EndStringIdentifier = $content | Where-Object{$_.Contains("--,,")} | Select-Object -Unique -First 1
$tableEnd = $content.IndexOf($EndStringIdentifier) - 1
$csv = $content[$tableStart..$tableEnd] | Out-File  "$workingDirectory\$reportDate.csv"
$data = Import-Csv -Path "$workingDirectory\$reportDate.csv"
#$data = Import-Csv -Path "DevOps_Capability_Assessment_Jan_19_2023_9_24_21_AM.csv"
$data | % { $_.Weight = [int]$_.Weight }  # fails if weight blank
$pillars = $data.Category | Select-Object -Unique


#region CSV Calculations

$agileDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Agile Software Development" -and $_.Category -eq "Survey Level Group"}).Description
$versionDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Version Control" -and $_.Category -eq "Survey Level Group"}).Description
$cicdDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Continuous Integration and Continuous Delivery (CI/CD)" -and $_.Category -eq "Survey Level Group"}).Description
$infraDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Infrastructure as a Flexible Resource" -and $_.Category -eq "Survey Level Group"}).Description
$securityDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Continuous Security" -and $_.Category -eq "Survey Level Group"}).Description
$configDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Configuration Management" -and $_.Category -eq "Survey Level Group"}).Description
$monitorDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq "Continuous Monitoring" -and $_.Category -eq "Survey Level Group"}).Description

function Get-PillarInfo($pillar)
{
    if($pillar.Contains("Agile Software Development"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $agileScore; "Description" = $agileDescription; "ScoreDescription" = $OverallScoreDescription}
    }
    if($pillar.Contains("Version Control"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $versionScore; "Description" = $versionDescription; "ScoreDescription" = $versionScoreDescription}
    }
    if($pillar.Contains("Continuous Integration and Continuous Delivery (CI/CD)"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $cicdScore; "Description" = $cicdDescription; "ScoreDescription" = $cicdScoreDescription}
    }
    if($pillar.Contains("Infrastructure as a Flexible Resource"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $infraScore; "Description" = $infraDescription; "ScoreDescription" = $infraScoreDescription}
    }
    if($pillar.Contains("Continuous Security"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $securityScore; "Description" = $securityDescription; "ScoreDescription" = $securityScoreDescription}
    }
    if($pillar.Contains("Continuous Monitoring"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $monitorScore; "Description" = $monitorDescription; "ScoreDescription" = $monitorScoreDescription}
    }
    if($pillar.Contains("Configuration Management"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $configScore; "Description" = $configDescription; "ScoreDescription" = $configScoreDescription}
    }
    if($pillar.Contains("Culture"))
    {
        return [pscustomobject]@{"Pillar" = $pillar; "Score" = $cultureScore; "Description" = $cultureDescription; "ScoreDescription" = $cultureScoreDescription}
    }
}

$overallScore = ""
$agileScore = ""
$versionScore = ""
$cicdScore = ""
$infraScore = ""
$securityScore = ""
$monitorScore = ""
$configScore = ""
$cultureScore = ""
$overallScoreDescription = ""
$agileScoreDescription = ""
$versionScoreDescription = ""
$cicdScoreDescription = ""
$infraScoreDescription = ""
$securityScoreDescription = ""
$monitorScoreDescription = ""
$configScoreDescription = ""
$cultureScoreDescription = ""

for($i=3; $i -le 8; $i++)
{
    if($Content[$i].Contains("overall"))
    {
        $overallScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $overallScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Agile Software Development"))
    {
        $agileScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $agileScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Version Control"))
    {
        $versionScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $versionScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Continuous Integration and Continuous Delivery (CI/CD)"))
    {
        $cicdScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $cicdScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Infrastructure as a Flexible Resource"))
    {
        $infraScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $infraScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Continuous Security"))
    {
        $securityScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $securityScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Continuous Monitoring"))
    {
        $monitorScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $monitorScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Configuration Management"))
    {
        $configScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $configScoreDescription = $Content[$i].Split(',')[1]
    }
    if($Content[$i].Contains("Culture"))
    {
        $cultureScore = $Content[$i].Split(',')[2].Trim("'").Split('/')[0]
        $cultureScoreDescription = $Content[$i].Split(',')[1]
    }
}

#endregion



#region Instantiate PowerPoint variables
Add-type -AssemblyName c:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll
$application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
$presentation = $application.Presentations.open($templatePresentation)
$titleSlide = $presentation.Slides[7]
$summarySlide = $presentation.Slides[8]
$detailSlide = $presentation.Slides[9]

#endregion

#region Clean the uncategorized data

if($data.PSobject.Properties.Name -contains "ReportingCategory"){
    foreach($lineData in $data)
    {
        
        if(!$lineData.ReportingCategory)
        {
            $lineData.ReportingCategory = "Uncategorized"
        }
    }
}

#endregion

foreach($pillar in $pillars)
{
 $pillarData = $data | Where-Object{$_.Category -eq $pillar}
 $pillarInfo = Get-PillarInfo -pillar $pillar
 # Edit title & date on slide 1
 $slideTitle = $title.Replace("[pillar]", $pillar) #,$pillar.substring(0,1).toupper()+$pillar.substring(1).tolower()) #lowercase only here?
 $newTitleSlide = $titleSlide.Duplicate()
 $newTitleSlide.MoveTo($presentation.Slides.Count)
 $newTitleSlide.Shapes[3].TextFrame.TextRange.Text = $slideTitle
 $newTitleSlide.Shapes[4].TextFrame.TextRange.Text = $newTitleSlide.Shapes[4].TextFrame.TextRange.Text.Replace("[Report_Date]",$localReportDate)

 # Edit Executive Summary Slide

 #Add logic to get overall score
 $newSummarySlide = $summarySlide.Duplicate()
 $newSummarySlide.MoveTo($presentation.Slides.Count)

 if(![string]::IsNullOrEmpty($pillarInfo.Score)){
     $ScoreText = "$($pillarInfo.Score) - $($pillarInfo.ScoreDescription)"
 }
 else{
    $ScoreText = "$($pillarInfo.ScoreDescription)"
 }
  
 $newSummarySlide.Shapes[4].TextFrame.TextRange.Text = $ScoreText
 $newSummarySlide.Shapes[5].TextFrame.TextRange.Text = $pillarInfo.Description

 $CategoriesList = New-Object System.Collections.ArrayList
 $categories = ($pillarData | Sort-Object -Property "Weight" -Descending).ReportingCategory | Select-Object -Unique
 foreach($category in $categories)
 {
    $categoryWeight = ($pillarData | Where-Object{$_.ReportingCategory -eq $category}).Weight | Measure-Object -Sum
    $categoryScore = $categoryWeight.Sum/$categoryWeight.Count
    $categoryWeightiestCount = ($pillarData | Where-Object{$_.ReportingCategory -eq $category}).Weight -ge $MinimumReportLevel | Measure-Object
    $CategoriesList.Add([pscustomobject]@{"Category" = $category; "CategoryScore" = $categoryScore; "CategoryWeightiestCount" = $categoryWeightiestCount.Count}) | Out-Null
 }

 $CategoriesList = $CategoriesList | Sort-Object -Property CategoryScore -Descending

 $counter = 9 #Shape count for the slide to start adding scores
 foreach($category in $CategoriesList)
 {
    if($category.Category -ne "Uncategorized")
    {
        try
        {
            #$newSummarySlide.Shapes[8] #Domain 1 Icon
            #$newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryScore.ToString("#")
            $newSummarySlide.Shapes[$counter].TextFrame.TextRange.Text = $category.CategoryWeightiestCount.ToString("#")
            $newSummarySlide.Shapes[$counter+1].TextFrame.TextRange.Text = $category.Category
            $counter = $counter + 2 # no graphic anymore
        }
        catch{}
    }
 }

 #Remove the boilerplate placeholder text if categories < 8
 if($categories.Count -lt 8)
 {
     for($k=$newSummarySlide.Shapes.count; $k -gt $counter-1; $k--)
     {
        try
        {
         $newSummarySlide.Shapes[$k].Delete()
         <#$newSummarySlide.Shapes[$k].Delete()
         $newSummarySlide.Shapes[$k+1].Delete()#>
         }
         catch{}
     }
 }

 # Edit new category summary slide

 foreach($category in $CategoriesList.Category)
 {
    $BlurbIndex=1
    $TitleIndex=2 
    $ScoreIndex = 5
    $DescriptionIndex = 6
    $InnerTitleIndex=9
    $ContentIndex=10

    $categoryData = $pillarData | Where-Object{$_.ReportingCategory -eq $category -and $_.Category -eq $pillar}
    $categoryDataCount = ($categoryData | measure).Count
    $categoryWeight = ($pillarData | Where-Object{$_.ReportingCategory -eq $category}).Weight | Measure-Object -Sum
    $categoryScore = $categoryWeight.Sum/$categoryWeight.Count
    $categoryDescription = ($descriptionsFile | Where-Object{$_.Pillar -eq $pillar -and $categoryData.ReportingCategory.Contains($_.Category)}).Description
    $y = $categoryDataCount
    $x = $ShowTop
    if($categoryDataCount -lt $x)
    {
        $x = $categoryDataCount
    }

    $newDetailSlide = $detailSlide.Duplicate()
    $newDetailSlide.MoveTo($presentation.Slides.Count)

    $newDetailSlide.Shapes[$TitleIndex].TextFrame.TextRange.Text = $category
    if($category -eq "Uncategorized"){
        $newDetailSlide.Shapes[$BlurbIndex].TextFrame.TextRange.Text = ""
        $newDetailSlide.Shapes[$ScoreIndex].TextFrame.TextRange.Text = ""
        $newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Text = ""
        $newDetailSlide.Shapes[$DescriptionIndex].TextFrame.TextRange.Text = "Uncategorized items are typically technical - for instance, from Azure Advisor - or aren't sourced from the Well-Architected Review survey directly.`r`n`r`nPlease refer to your Work Items list for the complete set."
    }
    else{
        $newDetailSlide.Shapes[$ScoreIndex].TextFrame.TextRange.Text = $categoryScore.ToString("#")
        $newDetailSlide.Shapes[$DescriptionIndex].TextFrame.TextRange.Text = $categoryDescription
    }
    $newDetailSlide.Shapes[$InnerTitleIndex].TextFrame.TextRange.Text = "Top $x of $y recommendations:"
    
    $newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Text = ($categoryData | Sort-Object -Property "Link-Text" -Unique | Sort-Object -Property Weight -Descending | Select-Object -First $x).'Link-Text' -join "`r`n`r`n"
    $sentenceCount = $newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Sentences().count
    
    for($k=1; $k -le $sentenceCount; $k++)
     {
         if($newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Sentences($k).Text)
         {
            try
            {
                $recommendationObject = $categoryData | Where-Object{$newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Sentences($k).Text.Contains($_.'Link-Text')}
                $newDetailSlide.Shapes[$ContentIndex].TextFrame.TextRange.Sentences($k).ActionSettings(1).HyperLink.Address = $recommendationObject.Link
            }
            catch{}
         }
     }    

 }

 }

 $titleSlide.Delete()
 $summarySlide.Delete()
 $detailSlide.Delete()

 $presentation.SavecopyAs("$workingDirectory\DOCA-Review-$($reportDate).pptx")
 $presentation.Close()


$application.quit()
$application = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()