# The sample scripts provided here are not supported under any Microsoft standard support program or service. All scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or performance of the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, even if Microsoft has been advised of the possibility of such damages.

# This script is intented for folders like Notes, Calendar or Tasks - to prevent their contents from getting archived by a mailbox-wide MRM Policy.

function StampPolicyOnFolder {
    param (
        [String] $EwsDll,
        [String] $TenantId,
        [String] $ApplicationId,
        [String] $Secret,
        [String] $MailboxName,
        [String] $FolderName,
        [String] $TagId
    )

    Add-Type -AssemblyName System.Web
    Add-Type -Path $EwsDll

    $LogFile = ".\StampPolicyOnFolder-ScriptLog.txt"

    if (!(Test-Path -Path $LogFile)) {
        New-Item -Path . -Name "StampPolicyOnFolder-ScriptLog.txt" -ItemType File | Out-Null
    }

    # To obtain an OAuth token
    $Body = @{
        client_id     = $ApplicationId
        client_secret = $Secret
        scope         = "https://protect-us.mimecast.com/s/NNUxCBBv85U2WZ67H3OdQ4?domain=outlook.office365.com"
        grant_type    = "client_credentials"
    }

    $PostSplat = @{
        ContentType = "application/x-www-form-urlencoded"
        Method      = "POST"
        Body        = $Body
        Uri         = "https://protect-us.mimecast.com/s/w-uvCDkxY5uGY6v5fJpWa_?domain=login.microsoftonline.com"
    }

    $Token = (Invoke-RestMethod @PostSplat).access_token

    # Create an EWS entry point
    $Service = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService

    $Service.Url = "https://protect-us.mimecast.com/s/Ofy5CNkGEquyknl0U06F-2?domain=outlook.office365.com"
    $Service.EnableScpLookup = $false

    # $Service.TraceEnabled = $true

    $Service.HttpHeaders.Add("X-AnchorMailbox", $MailboxName)

    # Change the user to Impersonate
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)

    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token)

    # Start applying the Tag

    $FolderFoundOrCreate = $false
    $CreateFolderIfDoesNotExist = $false

    Write-Host "Searching for $($MailboxName)" -ForegroundColor Yellow
    Add-Content -Path $LogFile -Value "$([DateTime]::Now) : Initiating script for `"$($MailboxName)`""

    $oFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)

    $oSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)

    $oFindFolderResults = $Service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot, $oSearchFilter, $oFolderView)

    if ($oFindFolderResults.TotalCount -eq 0) {
        Write-Host "Couldn't find a `"$($FolderName)`" folder in mailbox `"$($MailboxName)`"" -ForegroundColor Yellow
        Add-Content -Path $LogFile -Value "$([DateTime]::Now) : Folder `"$($FolderName)`" not found for `"$($MailboxName)`""

        # If the folder doesn't exist, create it:
        if ($CreateFolderIfDoesNotExist -eq $true) {
            $oFolder = New-Object Microsoft.Exchange.WebServices.Data.Folder($Service)
            $oFolder.DisplayName = $FolderName
            $oFolder.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot)

            $FolderFoundOrCreate = $true
            Write-Host "`"$($FolderName)`" folder created in mailbox `"$($MailboxName)`"" -ForegroundColor Yellow
            Add-Content -Path $LogFile -Value "$([DateTime]::Now) : Folder `"$($FolderName)`" created for `"$($MailboxName)`""
        }
    }
    else {
        Write-Host "`"$($FolderName)`" folder found in mailbox `"$($MailboxName)`"" -ForegroundColor Green
        Add-Content -Path $LogFile -Value "$([DateTime]::Now) : Folder `"$($FolderName)`" found for `"$($MailboxName)`""

        # Bind to the Folder
        $oFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $oFindFolderResults.Folders[0].Id)
        $FolderFoundOrCreate = $true
    }


    if ($FolderFoundOrCreate -eq $true) {
        # PR_ARCHIVE_TAG 0x3018
        $ArchiveTag = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3018, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)

        # PR_RETENTION_FLAGS 0x301D
        $RetentionFlags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
        
        # PR_ARCHIVE_PERIOD 0x301E
        $RetentionPeriod = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301E, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)

        # Same as the MRM policy
        $oFolder.SetExtendedProperty($RetentionFlags, 16)
        $oFolder.SetExtendedProperty($RetentionPeriod, 0)

        # Change the GUID based on your policy tag
        $ArchiveTagGUID = New-Object Guid($TagId)

        $oFolder.SetExtendedProperty($ArchiveTag, $ArchiveTagGUID.ToByteArray())

        $oFolder.Update()

        Write-Host "PRT tagged on the `"$($FolderName)`" folder in mailbox `"$($MailboxName)`"" -ForegroundColor Green
        Add-Content -Path $LogFile -Value "$([DateTime]::Now) : PRT $($TagId) stamped on the `"$($FolderName)`" folder for `"$($MailboxName)`"`n"
    }

    $Service.ImpersonatedUserId = $null

}
