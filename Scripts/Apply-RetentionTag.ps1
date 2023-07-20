# PRE_REQUISITES
<#
    # 1. Connect to Exchange Online to retrieve a "Never Delete" retention tag GUID
    Connect-ExchangeOnline
    $retentionTag = Get-RetentionPolicyTag "Never Delete" | Select-Object RetentionId

    # If a retention tag does not exist yet, we can create one, and apply it to an existing Retention Policy:
    <#
        New-RetentionPolicyTag `
            -Name "Never Delete" `
            -Type Personal `
            -Comment "Personal tag meant to prevent archival of items under the Calendars, Notes, and Tasks folders." `
            -RetentionEnabled $false ` # When set to $false, the tag is disabled, and no retention action is taken on messages that have the tag applied.
            -RetentionAction DeleteAndAllowRecovery # Deletes a message and allows recovery from the Recoverable Items folder.
    #>
#>

# Hardcoded here for simplicity, but the $retentionTag.RetentionId value can be passed below
$retentionIdBytes = (New-Object -TypeName Guid -ArgumentList "414c6a14-3ed5-432e-9edb-c6620a8278f0").ToByteArray()

# - Authenticate an EWS application by using OAuth
# https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

# Import the EWS Managed API DLL
Add-Type -Path "$($PSScriptRoot)\Dependencies\Microsoft.Exchange.WebServices.dll"

$applicationDetails = Import-Clixml -Path "$($PSScriptRoot)\Secrets\Azure AD App Registrations.xml"

. $PSScriptRoot\Get-OAuthToken.ps1

# Create the ExchangeService object
$ewsClient = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService
$ewsClient.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$ewsClient.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials `
    -ArgumentList (
        Get-OAuthToken -TenantId ($applicationDetails."EWS".TenantID) `
            -ClientId ($applicationDetails."EWS".AppID) `
            -ClientSecret ($applicationDetails."EWS".Secret) `
            -Scope ($applicationDetails."EWS".Scopes)
    )

# Set the $targetMailbox for the script
$targetMailbox = "Admin@luspin.onmicrosoft.com"

$ewsClient.HttpHeaders.Add("X-AnchorMailbox", $targetMailbox)
$ewsClient.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
    -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress), $targetMailbox

# Define the folder names to be stamped with the Retention Tag
$targetFolderNames = @("Calendar", "Notes", "Tasks")

foreach ($folderName in $targetFolderNames) {
    $folderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $folderName)
    $findFolderResults = $ewsClient.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $searchFilter, $folderView)
    
    if ($findFolderResults.TotalCount -eq 0) {
        Write-Host "Folder does not exist in Mailbox: $($targetMailbox)" -ForegroundColor Red
    }
    else {
        Write-Host "Folder found in Mailbox:" $targetMailbox -ForegroundColor Green
    
        # PR_ARCHIVE_TAG 0x3018 – We use the PR_ARCHIVE_TAG instead of the PR_POLICY_TAG
        $PR_ARCHIVE_TAG= New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
    
        # PR_RETENTION_FLAGS 0x301D
        $PR_RETENTION_FLAGS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

        # PR_ARCHIVE_PERIOD 0x301E - We use the PR_ARCHIVE_PERIOD instead of the PR_RETENTION_PERIOD 
        $PR_ARCHIVE_PERIOD = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
    
        # Bind to the $folderName folder
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $findFolderResults.Folders[0].Id)

        # Stamp the PR_ARCHIVE_TAG
        $folder.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes)

        # Stamp the PR_RETENTION_FLAGS - 16 specifies that this is a ExplictArchiveTag
        $folder.SetExtendedProperty($PR_RETENTION_FLAGS, 16)

        # Stamp the PR_ARCHIVE_PERIOD - Since this tag is disabled the Period would be 0
        $folder.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)

        # Save changes to the $folder
        $folder.Update()
    
        Write-Host "Retention policy stamped on the `"$($folderName)`" folder." -ForegroundColor Green

        # PART 2: Stamp the same tags on Items in the folder
        # Set the view to retrieve all items in the folder
        $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView([int]::MaxValue)

        # Perform the FindItems search in the folder
        $findItemsResults = $folder.FindItems($itemView)

        # Loop through the items and process them
        foreach ($item in $findItemsResults) {
            # Process each item as needed
            Write-Host "Subject: $($item.Subject)"

            # Stamp the PR_ARCHIVE_TAG
            $item.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes )

            # Stamp the PR_RETENTION_FLAGS - 16 specifies that this is a ExplictArchiveTag
            $item.SetExtendedProperty($PR_RETENTION_FLAGS, 16)
        
            # Stamp the PR_ARCHIVE_PERIOD - Since this tag is disabled the Period would be 0
            $item.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)
        
            # Save changes to the $item
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
        }
    }
}