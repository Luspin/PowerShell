# Connect to Exchange Online to retrieve a "Never Delete" retention tag GUID
# Connect-ExchangeOnline
# $retentionTag = Get-RetentionPolicyTag "Tag Name" | Select-Object RetentionId

# Hardcoded here
$retentionTag = "414c6a14-3ed5-432e-9edb-c6620a8278f0"
$archiveTagGuid = New-Object -TypeName Guid -ArgumentList $retentionTag
$guidBytes = $archiveTagGuid.ToByteArray()

# - Authenticate an EWS application by using OAuth
# https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth

# Import the EWS Managed API DLL
Add-Type -Path "$($PSSCriptRoot)\Microsoft.Exchange.WebServices.dll"

# Provide your Office 365 Tenant Domain Name or Tenant Id
$tenantId = "REDACTED.onmicrosoft.com"
# Provide Application (client) Id of your app
$appClientId = "832201e4-9f21-4652-8562-c7d1649c0d24"
# Provide Application client secret key
$clientSecret = "REDACTED"

# Retrieve an OAuth token using the client credentials grant flow
# https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
$requestBody = @{
    client_id     = $appClientId;
    client_secret = $clientSecret;
    grant_type    = "client_credentials";
    scope         = "https://outlook.office365.com/.default"
}

$OAuthResponse = Invoke-RestMethod `
                    -Method Post `
                    -Uri "https://login.microsoftonline.com/$($tenantId)/oauth2/v2.0/token" `
                    -Body $requestBody

$oAuthCredentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $OAuthResponse.access_token

# Create the ExchangeService object
$ewsClient = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService
$ewsClient.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$ewsClient.Credentials = $oAuthCredentials

$targetMailbox = "Admin@luspin.onmicrosoft.com"

$ewsClient.ImpersonatedUserId = New-Object `
                                    -TypeName Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
                                    -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress), $targetMailbox

$ewsClient.HttpHeaders.Add("X-AnchorMailbox", $targetMailbox);

# define the folder name to be stamped with the Retention Tag
$targetFolderNames = @("Calendar", "Notes", "Tasks")

foreach ($folderName in $targetFolderNames) {
    $oFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $oSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $folderName)
    $oFindFolderResults = $ewsClient.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $oSearchFilter, $oFolderView)
    
    if ($oFindFolderResults.TotalCount -eq 0) {
        Write-host "Folder does not exist in Mailbox: $($targetMailbox)" -foregroundcolor Red
        Add-Content $LogFile ("Folder does not exist in Mailbox: $($targetMailbox)")
    }
    else {
        Write-host "Folder found in Mailbox:" $targetMailbox -foregroundcolor Green
    
        #PR_ARCHIVE_TAG 0x3018 – We use the PR_ARCHIVE_TAG instead of the PR_POLICY_TAG
        $ArchiveTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
    
        #PR_RETENTION_FLAGS 0x301D
        $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
    
        #PR_ARCHIVE_PERIOD 0x301E - We use the PR_ARCHIVE_PERIOD instead of the PR_RETENTION_PERIOD
        $ArchivePeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
    
        #Bind to the folder found
        $oFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ewsClient, $oFindFolderResults.Folders[0].Id)
    
        #Stamp GUID
        $oFolder.SetExtendedProperty($ArchiveTag, $guidBytes )
    
        #Same as that on the policy - 16 specifies that this is a ExplictArchiveTag
        $oFolder.SetExtendedProperty($RetentionFlags, 16)
    
        #Same as that on the policy - Since this tag is disabled the Period would be 0
        $oFolder.SetExtendedProperty($ArchivePeriod, 0)
    
        $oFolder.Update()
    
        Write-host "Retention policy stamped!" -foregroundcolor Green
        # Add-Content $LogFile ("Retention policy stamped!")
}


}