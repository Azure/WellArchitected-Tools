param(    
    [Parameter()][string]
    $Branch = "main"
    )
#This script will download a list of files from the proper location to the directory it is running from

#This is the base URL for downloads. Base URL cannot end with a /
$baseURL = "https://raw.githubusercontent.com/Azure/WellArchitected-Tools/$Branch/WARP/devops"

$workingDirectory = (Get-Location).Path
Write-Host "Working Directory: $workingDirectory"
Invoke-WebRequest $baseURL/files-list.txt -OutFile $workingDirectory/files-list.txt


Write-Host "Downloading from: $baseURL"
Write-Host "We will get these files:"
Get-Content $workingDirectory/files-list.txt | ForEach-Object {Write-Host "   $_"}


Get-Content $workingDirectory/files-list.txt | ForEach-Object {Invoke-WebRequest $baseURL/$_ -OutFile $workingDirectory/$(Split-Path $_ -Leaf)}
