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
            $contact += ($FileData[$i] + "`r`n").TrimStart()
            $i++
        } until (
            $FileData[$i] -match "^(END:VCARD)"
        )

        $contact += "END:VCARD" + "`r`n"
        $contactsCollection += $contact
        $contact  = $null
    }

    return $contactsCollection

}

$i = 0

$contactsCollection = Split-Cards -FilePath $PSScriptRoot\Contacts.vcf

$contactsCollection | ForEach-Object {
    $_ | Out-File -FilePath "$PSScriptRoot\$($i).vcf" -Encoding UTF8;
    $i++
}