function Get-PidTagMessageClassProperties {

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Int]$Value
    )

    $Properties = @{
        mfRead         = @{ "IsSet" = (($Value -band 0x00000001) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_READ" }
        mfUnsent       = @{ "IsSet" = (($Value -band 0x00000008) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_UNSENT" }
        mfResend       = @{ "IsSet" = (($Value -band 0x00000080) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_RESEND" }
        mfUnmodified   = @{ "IsSet" = (($Value -band 0x00000002) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_UNMODIFIED" }
        mfSubmitted    = @{ "IsSet" = (($Value -band 0x00000004) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_SUBMIT" }
        mfHasAttach    = @{ "IsSet" = (($Value -band 0x00000010) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_HASATTACH" }
        mfFromMe       = @{ "IsSet" = (($Value -band 0x00000020) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_FROMME" }
        mfFAI          = @{ "IsSet" = (($Value -band 0x00000040) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_ASSOCIATED" }
        mfNotifyRead   = @{ "IsSet" = (($Value -band 0x00000100) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_RN_PENDING" }
        mfNotifyUnread = @{ "IsSet" = (($Value -band 0x00000200) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_NRN_PENDING" }
        mfEverRead     = @{ "IsSet" = (($Value -band 0x00000400) -ne 0); "PR_MESSAGE_FLAG" = "N/A" }
        mfInternet     = @{ "IsSet" = (($Value -band 0x00002000) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_ORIGIN_INTERNET" }
        mfUntrusted    = @{ "IsSet" = (($Value -band 0x00008000) -ne 0); "PR_MESSAGE_FLAG" = "MSGFLAG_ORIGIN_MISC_EXT" }
    }

    return $Properties.GetEnumerator() | Where-Object { $_.Value.IsSet }

}