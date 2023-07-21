function Split-Cards {
    param (
        [String] $FilePath
    )

    # Import the VCF file
    $FileData = (Get-Content $FilePath)

    # Find lines that start with "BEGIN:VCARD"
    $BeginVCardLines = $FileData -match "^(BEGIN:VCARD)" | Select-Object ReadCount

    # Instantiate a "$contactsCollection" array
    $contactsCollection = @()

    # Loop over each "BEGIN:VCARD" line
    foreach ($line in $BeginVCardLines) {
        $i = $line.ReadCount - 1

        do {
            if ($FileData[$i].StartsWith("FN:")) {
                $contactName = $FileData[$i]
            }

            $contact += ($FileData[$i] + "`r`n").TrimStart()
            $i++
        } until (
            $FileData[$i] -match "^(END:VCARD)"
        )

        $contact += "END:VCARD" + "`r`n"
        $contactsCollection += , @($contactName, [String[]]$contact)
        $contact = $null
    }

    return $contactsCollection

}

$i = 0

$contactsCollection = Split-Cards -FilePath "$($PSScriptRoot)\Contacts.vcf"

$contactsCollection | ForEach-Object {
    $contactName = $_[0].Replace("FN:", "")
    $_[1] | Out-File -FilePath "$PSScriptRoot\$($contactName).vcf" -Encoding UTF8
    $i++
}