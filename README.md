# Smartsheet-PowerShell-Backup
A simple script to backup sheets and attachments from Smartsheet. This was created out of necessity for a work project, however, I decided to add a few more neat features to it and throw it online. 

This PowerShell script is designed to interact with the Smartsheet API, allowing users to list, download, and manage sheets and attachments based on various parameters. The script supports multiple actions like listing all sheets, downloading sheets, fetching sheet details, and managing attachments.

## Features

    List all sheets: Retrieve a summary list of all sheets available.
    Download sheets: Download specific sheets by providing their unique ID.
    Get sheet details: Fetch detailed information about a specific sheet.
    Search locally: Filter sheets based on a search query.
    Manage attachments: Download and manage sheet attachments.

Getting Started
Prerequisites

    PowerShell 5.1 or higher
    An active Smartsheet account with API access

# Configuration
At the top of the script, you will see a few parameters.You must set your API Key and Output Directory.

    $apiToken = "YOUR API KEY"  # Replace with your Smartsheet API key
    $outputPath = "YOUR OUTPUT DIRECTORY"  # Path where sheets and attachments will be stored

### Optional Settings: You can adjust the script behavior using these optional settings.

    $noDownload = $false  # Set to $true to disable downloading. See "Automated Folder Management" for more information.
    $debug = $false  # Enable debug mode to display additional output
    $retentionMonths = 3  # Duration in months to keep downloaded files before deletion

### Get the details of a sheet

    Smartsheet -Action Get-Sheet -SheetID "2984863124639620"

### Download a sheet
    Smartsheet -Action Download-Sheet -SheetID "YOUR_SHEET_ID" -TargetDirectory $outputPath

### Get the details of an attachment
    Smartsheet -Action Get-Attachment -SheetID "YOUR_SHEET_ID"
    
### Download all attachments from a sheet
    Smartsheet -Action Download-Attachment -SheetID "YOUR_SHEET_ID" -TargetDirectory $outputPath
  
### To list all sheets
    Smartsheet -Action ListAll

### To search for a sheet:
    Smartsheet -Action SearchLocal -SearchQuery "Keystone"

# Automated Folder Management

The script includes functionality to manage folders based on the date and cleans up older backups based on the retention period set by the user. This is particularly useful for maintaining a clean and manageable file system. By default, the automatic folder management is turned off. 

#### Backup folder topography to be generated:
    Main target folder (whatever is chosen by the user)
        -Folder "current date" e.g. 2024-04-25
            -Folder SheetName
                -Sheet itself
                -SheetName_attachments folder (if applicable)
