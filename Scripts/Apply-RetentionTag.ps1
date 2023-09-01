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
$ewsClient = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
$ewsClient.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$ewsClient.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials `
    -ArgumentList (
        Get-OAuthToken -TenantId ($applicationDetails."EWS (Application)".TenantID) `
            -ClientId ($applicationDetails."EWS (Application)".AppID) `
            -ClientSecret ($applicationDetails."EWS (Application)".Secret) `
            -Scope ($applicationDetails."EWS (Application)".Scopes)
    )

# Set the $targetMailbox for the script
$targetMailbox = "Admin1@luspin.onmicrosoft.com"

$ewsClient.HttpHeaders.Add("X-AnchorMailbox", $targetMailbox)
$ewsClient.ImpersonatedUserId = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
    -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress), $targetMailbox

# https://gist.github.com/robderickson/ab87cfe11f1de5fe41654676273c8837

# Create objects for setting the MS-OXPROPS properties for the retention policy tag: PR_POLICY_TAG (0x3019), PR_RETENTION_FLAGS (0x301D), PR_RETENTION_PERIOD (0x301A)
## PR_POLICY_TAG: Binary; GUID of the tag you are applying Use Get-RetentionPolicyTag to get the GUID: Get-RetentionPolicyTag 'Tag Name Here' | Select-Object GUID. Reference: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/4e44a078-c129-45e1-8bf8-cc8026efdca0
## PR_RETENTION_FLAGS: Integer; Apply the tag to a folder and look up the flags with MFCMAPI. Reference: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/8e03d9d0-0d9d-4620-901c-2343747136eb
## PR_RETENTION_PERIOD: Integer; Matches the age limit in days of your tag. 0 never expires. Reference: https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/e2354980-35c0-4984-84b9-afbaf0ca1984
                

# PR_ARCHIVE_TAG 0x3018 – We use the PR_ARCHIVE_TAG instead of the PR_POLICY_TAG
$PR_ARCHIVE_TAG = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);

# PR_RETENTION_FLAGS 0x301D
$PR_RETENTION_FLAGS = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

# PR_ARCHIVE_PERIOD 0x301E - We use the PR_ARCHIVE_PERIOD instead of the PR_RETENTION_PERIOD 
$PR_ARCHIVE_PERIOD = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);

# Define the folder names to be stamped with the Retention Tag
$targetFolderNames = @("Calendar", "Notes", "Tasks")

foreach ($folderName in $targetFolderNames) {

    <#
    $oFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $oSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $folderName)
    $oFindFolderResults = $ewsClient.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $oSearchFilter, $oFolderView)
    
    if ($oFindFolderResults.TotalCount -eq 0) {
        Write-Host "Folder does not exist in Mailbox: $($targetMailbox)" -ForegroundColor Red
    }
    else {
        Write-Host "Folder `"$($folderName)`" found in Mailbox:" $targetMailbox -ForegroundColor Green
    #>

        # Bind to the $folderName folder, retrieve ALL properties
        $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, @($PR_RETENTION_FLAGS, $PR_ARCHIVE_TAG, $PR_ARCHIVE_PERIOD))
        $wellKnownFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$($folderName)
        $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $wellKnownFolderName, $PropertySet)
        # $folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $oFindFolderResults.Folders[0].Id, $PropertySet)

        $folder.ExtendedProperties

        # Stamp the PR_ARCHIVE_TAG
        $folder.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes)

        # Stamp the PR_RETENTION_FLAGS - 16 specifies that this is a ExplictArchiveTag
        $folder.SetExtendedProperty($PR_RETENTION_FLAGS, 16)

        # Stamp the PR_ARCHIVE_PERIOD - Since this tag is disabled the Period would be 0
        $folder.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)

        # Save changes to the $folder
        $folder.Update()
    
        Write-Host "Retention policy stamped on the `"$($folderName)`" folder." -ForegroundColor Green

        $folder.ExtendedProperties

        # TO CLEAR THE VALUE:
        $folder.RemoveExtendedProperty($PR_ARCHIVE_TAG) # working!
        # $folder.RemoveExtendedProperty($PR_RETENTION_FLAGS)  
        $folder.SetExtendedProperty($PR_RETENTION_FLAGS, 0x00000080)
        $folder.RemoveExtendedProperty($PR_ARCHIVE_PERIOD)  

        # Save changes to the $folder
        $folder.Update()

        Write-Host "Retention policy REMOVED on the `"$($folderName)`" folder." -ForegroundColor Green

        $folder.ExtendedProperties

        <#
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
            $item.SetExtendedProperty($PR_ARCHIVE_TAG, $retentionIdBytes)

            # Stamp the PR_RETENTION_FLAGS - 16 specifies that this is a ExplictArchiveTag
            $item.SetExtendedProperty($PR_RETENTION_FLAGS, 16)

            # Stamp the PR_ARCHIVE_PERIOD - Since this tag is disabled the Period would be 0
            $item.SetExtendedProperty($PR_ARCHIVE_PERIOD, 0)

            # Save changes to the $item
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)

            Read-Host

            # TO CLEAR THE VALUES
            $item.RemoveExtendedProperty($PR_ARCHIVE_TAG)
            $item.RemoveExtendedProperty($PR_RETENTION_FLAGS)
            $item.RemoveExtendedProperty($PR_ARCHIVE_PERIOD)

            # Save changes to the $item
            $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)


        }
        #>
    }
# }