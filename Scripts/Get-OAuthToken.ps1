# Microsoft identity platform and the OAuth 2.0 client credentials flow
# - First case: Access token request with a shared secret
# https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow

function Get-OAuthToken {

    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(
            Mandatory = $true,
            Position = 0)]
        [String] $TenantId,

        # 'Default' ParameterSet
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Default',
            Position = 1)]
        [String] $ClientId,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Default',
            Position = 2)]
        [String] $ClientSecret,

        [Parameter(
            Mandatory = $false,
            ParameterSetName = 'Default',
            Position = 3)]
        [String] $Scope,

        # 'Alternate' ParameterSet
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Alternate',
            Position = 1,
            ValueFromPipeline)]
        [Hashtable] $Options
    )

    # Build the $requestBody
    Write-Verbose -Message "Parsing the arguments and building the $requestBody"

    $requestBody = @{
        client_id     = $Options ? $Options.client_id : $ClientId;
        client_secret = $Options ? $Options.client_secret : $ClientSecret;
        scope         = $Scope ? $Scope : "https://graph.microsoft.com/.default";
        grant_type    = "client_credentials"
    }

    # Issue the $tokenRequest
    try {
        Write-Verbose -Message "Issuing the $tokenRequest"

        $tokenRequest = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
                            -Method POST `
                            -Headers @{"Content-Type" = "application/x-www-form-urlencoded" } `
                            -Body $requestBody

        return $tokenRequest.access_token

    }
    catch [System.Net.WebException] {
        Write-Output "Unable to retrieve an OAuth token:`n'$($_.Exception.Message)'"

        Exit
    }
}