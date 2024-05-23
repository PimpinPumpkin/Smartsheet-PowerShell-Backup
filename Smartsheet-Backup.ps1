#Version 0.5
#Copyright Pimpinpumpkin 2024

#Set your variables
$apiToken = "YOUR TOKEN HERE"
$outputPath = "YOUR OUTPUT DIRECTORY HERE (NTFS or POSIX path)"

#Set this to true if you don't want any of the folders/downloads to run
$noDownload = $true
$debug = $false
$retentionMonths = 3
$throttleLimit = 10
$theMinimumPSVersion = "7.4"

#Create an empty hashtable for any and all errors
$myErrors = @()

#Create log file path
$logFile = Join-Path -Path $outputPath -ChildPath "Smartsheet-Log-$(Get-Date -Format "yyyy-MM-dd").txt"

function Get-PowerShellVersion {
    try {
        $requiredVersion = [Version]$theMinimumPSVersion
        $currentVersion = $PSVersionTable.PSVersion

        if ($currentVersion -lt $requiredVersion) {
            #Write-Error "PowerShell $requiredVersion or higher is required. You are running $currentVersion."
            return 0
        } else {
            Write-Host "Powershell version is equal to or greater than $theMinimumPSVersion"
            return 1
        }
    } 
    catch {
        Write-Error "Error checking PowerShell version: $_"
        return 0
    }
}

$powershellVersionState = Get-PowerShellVersion

#Make sure we're running a valid Powershell version
if ($powershellVersionState -eq 0) {
    Write-Error "Error: Powershell version is less than $theMinimumPSVersion"
    exit
}

function Set-LegalName {
    param (
        [string]$Name
    )

    #Define illegal characters for Windows filenames
    $illegalCharsPattern = '[\\/:*?"<>|\x00-\x1F]+'

    #Check for and replace illegal characters
    $sanitizedName = $Name -replace $illegalCharsPattern, '-'

    #Remove leading and trailing whitespace
    $sanitizedName = $sanitizedName.Trim()

    #Remove leading and trailing periods
    $sanitizedName = $sanitizedName -replace '^\.*|\.*$', ''

    #Check for reserved names
    $reservedNames = @("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
    if ($reservedNames -contains $sanitizedName.ToUpper()) {
        $sanitizedName = "InvalidName-$sanitizedName"
    }

    #Output changes for debugging
    if ($sanitizedName -ne $Name) {
        Write-Host "Illegal characters were found and replaced:`n$sanitizedName"
        Write-Host "`nOriginal name:`n$Name"
    } else {
        $sanitizedName = $Name
    }

    return $sanitizedName
}

#Define Set-LegalName as a string because Powershell sucks and we need it for calling inside of a multi-threaded Foreach-Object
$psSucks = ${function:Set-LegalName}.ToString()

function verifySheetID {
    param(
        [string]$SheetID,
        [string]$errorText
    )
    if (-not $SheetID) {
        $theError = "Error: $errorText"
        Write-Error $theError
        $myErrors += $theError
        exit
    }
}

function Set-FolderExists {
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
        [string]$parentSheetID 
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

#Start logging
Start-Transcript -Path $logFile -Append

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

    #Confirm we have an API token
    if (-not $apiToken) {
        $theError = "Error: API token is not set. Please configure your Smartsheet API token in the environment variables."
        Write-Error $theError
        $myErrors += $theError
        return
    }

    #Build headers
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

                #Clean sheet names of illegal characters
                $sheetName = Set-LegalName -Name $theQuery.name

                #Append .xlsx extension
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
                    $theError = "Error: Failed to obtain the sheet name from the API theQuery. Make sure the sheet actually exists!"
                    Write-Error $theError
                    $myErrors += $theError
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
                        #Grab Set-LegalName as a string because Powershell is a horrible language
                        ${function:Set-LegalName} = $using:psSucks
                        $currentAttachment = $_
                        $sanitizedAttachmentName = Set-LegalName -Name $currentAttachment.Name
                        $fileName = Join-Path -Path $using:TargetDirectory -ChildPath $sanitizedAttachmentName
                        $downloadUri = $currentAttachment.AttachmentFirstURL
                        $theFileError = $false

                        #Debug information
                        if ($debug) {
                            Write-Host "First stage URL: $downloadUri"
                        }

                        #Wrap the last URLs call in a try block (not the prettiest thing but it works)...
                        try {
                            #Get the new URLs by navigating from the urlsFirstStage object URLs 
                            $attachmentNewURLS = Invoke-RestMethod -Uri $downloadUri -Method Get -Headers $using:headers
                            $theFileError = $false
                        } 
                        catch {
                            #Retry 1
                            Write-Host "Last URL download for $sanitizedAttachmentName failed, retrying after 3 seconds"
                            #...if there's an error, let's take a rest. Note, this only affects the current thread, but if more errors pile up, each subsequent failing thread run into the same throttle loop
                            Start-Sleep -Seconds 3
                            try {
                                #Try again
                                $attachmentNewURLS = Invoke-RestMethod -Uri $downloadUri -Method Get -Headers $using:headers
                            }
                            catch {
                                #Retry 2
                                Write-Host "Last URL download for $sanitizedAttachmentName failed, retrying after 6 seconds"
                                #If it fails after a 3 second break, let's rest for a little longer
                                Start-Sleep -Seconds 6
                                try {
                                    #Try again
                                    $attachmentNewURLS = Invoke-RestMethod -Uri $downloadUri -Method Get -Headers $using:headers
                                }
                                catch {
                                    #Give up
                                    Write-Host "Last URL download for $sanitizedAttachmentName failed, giving up"
                                    #Ok now we can blow up and throw an error
                                    $theError = "Error: Failed to download attachments: $($_.Exception.Message)"
                                    Write-Error $theError
                                    $myErrors += $theError
                                    $theFileError = $true
                                }
                            }
                        }
                        
                        #Debug information
                        if ($debug) {
                            try {
                                Write-Host "Last stage URL: $($attachmentNewURLS.url)"
                            }
                            catch {
                                $theError = "Error: Error getting last stage URL"
                                Write-Error $theError
                                $myErrors += $theError
                            }
                        }

                        #Conditional logic here to make it only try and print/download from the second stage URL if we have it
                        if ($theFileError -eq $false) {
                            #Actually download the new files from the final stage URLs (parallel processing speeds this up considerably)
                            if ($debug) {
                                Write-Host "Preparing to download $sanitizedAttachmentName"
                            }

                            try {
                                Invoke-RestMethod -Uri $attachmentNewURLS.url -Method Get -OutFile $fileName
                                Write-Host "Downloaded $sanitizedAttachmentName"
                            } catch {
                                #Who knows
                            }
                        }
                        elseif ($theFileError -eq $true) {
                            $theError = "Error: Some issue with the second stage URLs not working right."
                            Write-Error $theError
                            $myErrors += $theError
                        }

                    } -ThrottleLimit $throttleLimit
                }
                catch {
                    $theError = "Error: Failed to download attachments: $($_.Exception.Message)"
                    Write-Error $theError
                    $myErrors += $theError
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
        $theError = "Error: Failed to process the request: $($_.Exception.Message)"
        Write-Error $theError
        $myErrors += $theError
    }
}

#From here on out is the automatic backup portion! If you want to turn it on, you can set $noDownload to $false at the top of the script.
if ($noDownload -eq $false) {

    #Get the current date and format it for folder titles
    $currentDate = Get-Date -Format "yyyy-MM-dd"

    #Combine the base path with the new folder name based on the current date
    $newFolderPath = Join-Path -Path $outputPath -ChildPath $currentDate

    #Check if the path exists, if not, create it
    Set-FolderExists -FolderPath $newFolderPath

    #Grab a list of all sheets
    $allSheets = Smartsheet -Action ListAll

    #Iterate through said list of all sheets
    $allSheets.data | ForEach-Object {
        $currentOperator = $($_)

        #Brief directory topography to be generated:
        #Main target folder (whatever is predefined)
        #..Folder named current date
        #...Workspace (New)
        #....Sheetname folder
        #......Sheet itself
        #......Sheet_attachments folder

        #Query to get the workspace for the current sheet
        $uncleanWorkspaceName = (Smartsheet -Action Get-Sheet -SheetID $currentOperator.id).workspace.name

        #Sanitize the workspace name
        $workspaceName = Set-LegalName -Name $uncleanWorkspaceName

        #Add the workspace property to the current operator object only
        $currentOperator | Add-Member -MemberType NoteProperty -Name workspace -Value $workspaceName -Force
        
        #Build folder named current date with workspaces under it
        $sheetDownloadFolderPath = Join-Path -Path $outputPath -ChildPath $currentDate -AdditionalChildPath $workspaceName, $currentOperator.name
        Set-FolderExists -FolderPath $sheetDownloadFolderPath

        #Download each sheet to the workspace folder under the new timestamped folder
        Smartsheet -Action Download-Sheet -SheetID $currentOperator.id -TargetDirectory $sheetDownloadFolderPath

        #Check if we have attachments for a sheet, and download them if so
        $doWeHaveAttachments = (Smartsheet -Action Get-Attachment -SheetID $currentOperator.id).count
        if ($doWeHaveAttachments -ge 1 ) {
            #Build sheet attachments folder
            $sheetAttachmentFolder = Join-Path -Path $sheetDownloadFolderPath -ChildPath "$($currentOperator.name)_attachments"
            Set-FolderExists -FolderPath $sheetAttachmentFolder

            try {
                #Download the attachments to the target folder
                Smartsheet -Action Download-Attachment -SheetID $currentOperator.id -TargetDirectory $sheetAttachmentFolder
            }
            catch {
                Write-Error "Error: Issue detected: $_"
                $myErrors += "Error: Issue detected: $_"
            }
        }
    }

    #Clean up backups older than x months

    #Get the current date minus x months
    $timeAgo = (Get-Date).AddMonths(-$retentionMonths)

    #Create holder array for folder dates
    $myFolderDates = @()

    #Enumerate each folder in the directory
    Get-ChildItem -Path $outputPath -Directory | ForEach-Object {
        #Try to parse the folder name as a date
        $folderDate = $_.Name
        Write-Host "Found folder: $($_.Name)"
        try {
            $parsedDate = [DateTime]::ParseExact($folderDate, "yyyy-MM-dd", $null)

            #Add to our array of dates
            $myFolderDates += $parsedDate

            #If the parsed date is older than three months ago, delete the folder
            if ($parsedDate -lt $timeAgo) {
                Remove-Item -Path $_.FullName -Recurse -Force
                Write-Host "Deleted folder: $($_.FullName)"
            }
        }
        catch {
            Write-Host "Skipping: $folderDate is not in the 'yyyy-MM-dd' format."
        }
    }
}

Stop-Transcript


