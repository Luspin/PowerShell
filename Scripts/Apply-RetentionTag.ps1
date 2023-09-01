# PRE-REQUISITES
<#
    # 1. Connect to Exchange Online to retrieve a "Never Delete" retention tag GUID
    [PS] > Connect-ExchangeOnline
    [PS] > $retentionTag = Get-RetentionPolicyTag "Never Delete" | Select-Object RetentionId

    # 2. If a retention tag doesn't exist yet, we can create one, and apply it to an existing Retention Policy:
    [PS] > New-RetentionPolicyTag `
            -Name "Never Delete" `
            -Type Personal `
            -Comment "Personal tag meant to prevent archival of items under the Calendars, Notes, and Tasks folders." `
            -RetentionEnabled $false ` # When set to $false, the tag is disabled, and no retention action is taken on messages that have the tag applied.
            -RetentionAction DeleteAndAllowRecovery # Deletes a message and allows recovery from the Recoverable Items folder.
#>

# REFERENCES
# - Authenticate an EWS application by using OAuth
# https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

# robderickson/Add-RetentionTagToFolder.ps1
# https://gist.github.com/robderickson/ab87cfe11f1de5fe41654676273c8837

# Hardcoded for simplicity, but the $retentionTag.RetentionId value should be passed here:
$retentionIdBytes = (New-Object -TypeName Guid -ArgumentList "414c6a14-3ed5-432e-9edb-c6620a8278f0").ToByteArray()

# Import the EWS Managed API DLL
Add-Type -Path "$($PSScriptRoot)\Dependencies\Microsoft.Exchange.WebServices.dll"

# Import the Azure AD App Registration details
$applicationDetails = Import-Clixml -Path "$($PSScriptRoot)\Secrets\Azure AD App Registrations.xml"

. $PSScriptRoot\Get-OAuthToken.ps1

# Create the ExchangeService object
$ewsClient = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$ewsClient.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$ewsClient.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials `
    -ArgumentList (
        Get-OAuthToken `
            -TenantId ($applicationDetails."EWS (Application)".TenantID) `
            -ClientId ($applicationDetails."EWS (Application)".AppID) `
            -ClientSecret ($applicationDetails."EWS (Application)".Secret) `
            -Scope ($applicationDetails."EWS (Application)".Scopes)
    )

# Set the $targetMailbox for the script
$targetMailbox = "Admin1@luspin.onmicrosoft.com"

# Define the folder names to be stamped with the Retention Tag
$targetFolderNames = @("Calendar", "Notes", "Tasks")

$ewsClient.HttpHeaders.Add("X-AnchorMailbox", $targetMailbox)
$ewsClient.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
    -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress), $targetMailbox

# Create objects for setting the MS-OXPROPS properties for the Retention Policy Tag:
# PR_ARCHIVE_TAG      | 0x3018 | Binary  | 2.2.1.58.1 PidTagArchiveTag Property <https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/ce4c9c11-8cb6-4b6f-83e2-7c7696e5d0a1>
# PR_RETENTION_FLAGS  | 0x301D | Integer | 2.2.1.58.6 PidTagRetentionFlags Property <https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/8e03d9d0-0d9d-4620-901c-2343747136eb>
# PR_RETENTION_PERIOD | 0x301A | Integer | 2.2.1.58.7 PidTagArchivePeriod Property <https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/94296c17-f1f8-4cdb-9c2b-f088535058e7>

$PR_ARCHIVE_TAG = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
$PR_RETENTION_FLAGS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
$PR_ARCHIVE_PERIOD = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

foreach ($folderName in $targetFolderNames) {
    # PART 1: Manipulate the Retention Policy Tag on the folder itself
    # Bind to the $folderName folder, retrieve ALL properties
    $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, @($PR_RETENTION_FLAGS, $PR_ARCHIVE_TAG, $PR_ARCHIVE_PERIOD))
    $wellKnownFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$($folderName)
    $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $wellKnownFolderName, $PropertySet)

    # Stamp the PR_ARCHIVE_TAG
    $folder.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes)

    # Stamp the PR_RETENTION_FLAGS - 16 | 0x00000010 specifies that this is an "ExplictArchiveTag"
    $folder.SetExtendedProperty($PR_RETENTION_FLAGS, 16)

    # Stamp the PR_ARCHIVE_PERIOD  - Since this tag is disabled, the Period will be 0
    $folder.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)

    # Save changes to the $folder
    $folder.Update()

    Write-Host "Retention policy STAMPED on the `"$($folderName)`" folder." -ForegroundColor Green

    $folder.ExtendedProperties

    # TO CLEAR THE VALUES
    $folder.RemoveExtendedProperty($PR_ARCHIVE_TAG)
    $folder.SetExtendedProperty($PR_RETENTION_FLAGS, 0x00000080) # "NeedsRescan"
    $folder.RemoveExtendedProperty($PR_ARCHIVE_PERIOD)

    # Again, save changes to the $folder
    $folder.Update()

    Write-Host "Retention policy REMOVED on the `"$($folderName)`" folder." -ForegroundColor Green

    $folder.ExtendedProperties

    # PART 2: Manipulate the same tags on Items within the folder
    # Set the view to retrieve all items in the folder
    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView([int]::MaxValue)

    # Perform the FindItems search in the folder
    $findItemsResults = $folder.FindItems($itemView)

    # Loop through the items and process them
    foreach ($item in $findItemsResults) {
        Write-Host "Subject: $($item.Subject)"

        # Stamp the PR_ARCHIVE_TAG
        $item.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes)

        # Stamp the PR_RETENTION_FLAGS - 16 | 0x00000010 specifies that this is an "ExplictArchiveTag"
        $item.SetExtendedProperty($PR_RETENTION_FLAGS, 16)

        # Stamp the PR_ARCHIVE_PERIOD  - Since this tag is disabled, the Period will be 0
        $item.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)

        # Save changes to the $item
        $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)

        # TO CLEAR THE VALUES
        $item.RemoveExtendedProperty($PR_ARCHIVE_TAG)
        $item.SetExtendedProperty($PR_RETENTION_FLAGS, 0x00000080) # "NeedsRescan"
        $item.RemoveExtendedProperty($PR_ARCHIVE_PERIOD)

        # Again, save changes to the $item
        $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
    }
}