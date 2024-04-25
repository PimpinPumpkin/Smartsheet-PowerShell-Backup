#Set your variables
$apiToken = "YOUR TOKEN HERE"
$outputPath = "YOUR OUTPUT DIRECTORY HERE"

#Set this to true if you don't want any of the folders/downloads to run
$noDownload = $false
$debug = $false
$retentionMonths = 3

function verifySheetID {
    param(
        [string]$SheetID,
        [string]$errorText
    )
    if (-not $SheetID) {
        Write-Error $errorText
        exit
    }
}

function Ensure-FolderExists {
    param (
        [Parameter(Mandatory = $true)]
        [string]$FolderPath
    )

    #Check if the folder already exists
    if (-Not (Test-Path -Path $FolderPath)) {
        #Folder does not exist, so create it
        New-Item -Path $FolderPath -ItemType Directory
        Write-Host "Folder created: $FolderPath"
    }
    else {
        Write-Host "Folder already exists: $FolderPath"
    }
}

function attachmentObjectFirstURL {
    param (
        [object]$listOfAttachments,
        [string]$parentSheetID  #Add parameter for sheetURL
    )

    $attachmentList = @()
    foreach ($attachment in $listOfAttachments.data) {
        $attachmentList += [PSCustomObject]@{
            ID                 = $attachment.id
            Name               = $attachment.name
            SizeInKb           = $attachment.sizeInKb
            ParentSheetID      = $parentSheetID  #Add the sheetURL to each object
            AttachmentFirstURL = "https://api.smartsheet.com/2.0/sheets/$parentSheetID/attachments/$($attachment.id)"
        }
    }
    return $attachmentList
}

function Smartsheet {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("ListAll", "Download-Sheet", "Get-Sheet", "SearchLocal", "Download-Attachment", "Get-Attachment")]
        [string]$Action = "ListAll",

        [Parameter(Mandatory = $true, ParameterSetName = 'Download-Sheet')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Get-Sheet')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Download-Attachment')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Get-Attachment')]
        [string]$SheetID,

        [Parameter(Mandatory = $false, ParameterSetName = 'SearchLocal')]
        [string]$SearchQuery,

        [Parameter(Mandatory = $true, ParameterSetName = 'Download-Sheet')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Download-Attachment')]
        [string]$TargetDirectory,

        #This is literally a bug in my Powershell interpreter. If I delete these, everything breaks.
        [Parameter(Mandatory = $true, ParameterSetName = 'Download-Attachment')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Get-Attachment')]
        [string]$attachmentID
        ####
    )

    if (-not $apiToken) {
        Write-Error "API token is not set. Please configure your Smartsheet API token in the environment variables."
        return
    }

    $baseUri = "https://api.smartsheet.com/2.0/sheets"
    #$basedUri = "https://api.smartsheet.com/2.0/sheets/?includeAll=true"
    $headers = @{
        "Authorization" = "Bearer $apiToken"
    }

    switch ($Action) {
        "ListAll" {
            $uri = "$baseUri/?includeAll=true"
        }
        "Download-Sheet" {
            verifySheetID -SheetID $SheetID -errorText "SheetID is required for downloading a sheet."
            $uri = "$baseUri/$SheetID"
        }
        "SearchLocal" {
            $uri = "$baseUri/?includeAll=true"
        }
        "Download-Attachment" {
            verifySheetID -SheetID $SheetID -errorText "SheetID is required to download attachments from a sheet."
            $uri = "$baseUri/$SheetID/attachments/?includeAll=true"
        }
        "Get-Sheet" {
            verifySheetID -SheetID $SheetID -errorText "SheetID is required to fetch sheet details."
            $uri = "$baseUri/$SheetID"
        }
        "Get-Attachment" {
            verifySheetID -SheetID $SheetID -errorText "SheetID is required to fetch attachment details from a sheet."
            $uri = "$baseUri/$SheetID/attachments/?includeAll=true"
        }
    }

    try {
        $theQuery = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
        switch ($Action) {
            "ListAll" {
                $theQuery | Add-Member -MemberType NoteProperty -Name powershellQueryURI -Value $uri
                $theQuery
            }
            "Download-Sheet" {
                #Set the headers to add the xlsx mimeType
                $headers["Accept"] = "application/vnd.ms-excel"

                $sheetName = "$($theQuery.name).xlsx"

                if ($sheetName) {
                    #Join the destination directory and the filename
                    $fileName = Join-Path -Path $TargetDirectory -ChildPath $sheetName

                    #Debug information
                    if ($debug) {
                        Write-Host "PowershellQueryURI: $uri"
                    }

                    #Download the actual file
                    Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -OutFile $fileName
                    Write-Host "Downloaded sheet $sheetName saved to $TargetDirectory"
                }
                else {
                    Write-Error "Failed to obtain the sheet name from the API theQuery. Make sure the sheet actually exists!"
                }
            }
            Get-Sheet {
                $theQuery | Add-Member -MemberType NoteProperty -Name powershellQueryURI -Value $uri
                $theQuery
            }
            "SearchLocal" {
                #Debug information
                if ($debug) {
                    Write-Host "powershellQueryURI: $uri"
                }
                $theQuery.data | Where-Object { $_.name -match $SearchQuery }
            }
            "Download-Attachment" {       
                verifySheetID -SheetID $SheetID -errorText "SheetID is required to download attachments from a sheet."
                try {
                    #Grab an object containing the first stage URLs and file names
                    $urlsFirstStage = attachmentObjectFirstURL -listOfAttachments $theQuery -parentSheetID $SheetID
            
                    #Iterate over each attachment in parallel
                    $urlsFirstStage | ForEach-Object -Parallel {
                        $currentAttachment = $_
                        $fileName = Join-Path -Path $using:TargetDirectory -ChildPath $currentAttachment.Name
                        $downloadUri = $currentAttachment.AttachmentFirstURL
                        
                        #Debug information
                        if ($debug) {
                            Write-Host "First stage URL: $downloadUri"
                        }

                        #Get the new URLs by navigating from the urlsFirstStage object URLs 
                        $attachmentNewURLS = Invoke-RestMethod -Uri $downloadUri -Method Get -Headers $using:headers

                        #Actually download the new files from the final stage URLs (parallel processing speeds this up considerably)
                        Write-Host "Preparing to download $($currentAttachment.Name)"

                        #Debug information
                        if ($debug) {
                            Write-Host "Last stage URL: $($attachmentNewURLS.url)"
                        }
                        Invoke-RestMethod -Uri $attachmentNewURLS.url -Method Get -OutFile $fileName
                        Write-Host "Downloaded $($currentAttachment.Name)"

                    } -ThrottleLimit 10
                }
                catch {
                    Write-Error "Failed to download attachments: $($_.Exception.Message)"
                }
            }
            
            "Get-Attachment" {
                #Verify the sheet ID is provided; if not, the script will error out and stop execution.
                verifySheetID -SheetID $SheetID -errorText "SheetID is required to download attachments from a sheet."
                attachmentObjectFirstURL -listOfAttachments $theQuery -parentSheetID $SheetID
                #Debug information
                if ($debug) {
                    Write-Host "powershellQueryURI: $uri"
                    Write-Host
                }
            }
            Default {
                $theQuery | Add-Member -MemberType NoteProperty -Name powershellQueryURI -Value $uri
                $theQuery
            }
        }
    }
    catch {
        Write-Error "Failed to process the request: $($_.Exception.Message)"
    }
}

#This portion is what actually downloads everything
if ($noDownload -eq $false) {

    #Get the current date and format it for folder titles
    $currentDate = Get-Date -Format "yyyy-MM-dd"

    #Combine the base path with the new folder name based on the current date
    $newFolderPath = Join-Path -Path $outputPath -ChildPath $currentDate

    #Check if the path exists, if not, create it
    Ensure-FolderExists -FolderPath $newFolderPath

    #Grab a list of all sheets
    $allSheets = Smartsheet -Action ListAll

    #Iterate through said list of all sheets
    $allSheets.data | ForEach-Object {
        $currentOperator = $($_)

        #Brief directory topography to be generated:
        #Main target folder (whatever is predefined)
        #..Folder named current date
        #....Sheetname folder
        #......Sheet itself
        #......Sheet_attachments folder

        #Build folder named current date
        $sheetDownloadFolderPath = Join-Path -Path (Join-Path -Path $outputPath -ChildPath $currentDate) -ChildPath $currentOperator.name
        Ensure-FolderExists -FolderPath $sheetDownloadFolderPath

        #Download each sheet to the new timestamped folder
        Smartsheet -Action Download-Sheet -SheetID $currentOperator.id -TargetDirectory $sheetDownloadFolderPath

        #Check if we have attachments for a sheet, and download them if so
        $doWeHaveAttachments = (Smartsheet -Action Get-Attachment -SheetID $currentOperator.id).count
        if ($doWeHaveAttachments -ge 1 ) {
            #Build sheet attachments folder
            $sheetAttachmentFolder = Join-Path -Path $sheetDownloadFolderPath -ChildPath "$($currentOperator.name)_attachments"
            Ensure-FolderExists -FolderPath $sheetAttachmentFolder

            #Download the attachments to the target folder
            Smartsheet -Action Download-Attachment -SheetID $currentOperator.id -TargetDirectory $sheetAttachmentFolder
        }
    }

    #Clean up backups older than x months (Acronis backups will have older versions)

    #Get the current date minus x months
    $timeAgo = (Get-Date).AddMonths(-$retentionMonths)

    #Enumerate each folder in the directory
    Get-ChildItem -Path $outputPath -Directory | ForEach-Object {
        #Try to parse the folder name as a date
        $folderDate = $_.Name
        Write-Host "Found folder: $($_.Name)"
        try {
            $parsedDate = [DateTime]::ParseExact($folderDate, "yyyy-MM-dd", $null)

            #If the parsed date is older than three months ago, delete the folder
            if ($parsedDate -lt $timeAgo) {
                Remove-Item -Path $_.FullName -Recurse -Force
                Write-Host "Deleted folder: $($_.FullName)"
            }
        }
        catch {
            Write-Host "Skipping: $folderDate is not in the 'yyyy-MM-dd' format or another error occurred."
        }
    }
}
