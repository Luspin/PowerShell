# STEP 1: Get the HTML file
# How to create an open file and folder dialog box with PowerShell
# https://github.com/myusefulrepo/Tips/blob/master/Tips%20%20-%20How%20to%20create%20an%20open%20file%20and%20folder%20dialog%20box%20with%20PowerShell.md
Add-Type -AssemblyName "System.Windows.Forms"

$FileBrowser = New-Object -TypeName System.Windows.Forms.OpenFileDialog `
    -Property @{ Filter = 'HTML Source File (*.html)|*.html' }

$null = $FileBrowser.ShowDialog()

$messageBody = Get-Content -Path $FileBrowser.FileNames -Raw

# STEP 2: Credentials
$credential = Get-Credential

# STEP 3: Relay the HTML message
# about Splatting - PowerShell | Microsoft Docs
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_splatting?view=powershell-7.2

$parameters = @{
    SmtpServer = "smtp.office365.com";
    Port       = 587;
    UseSsl     = $true;
    Credential = $credential;
    From       = $credential.UserName;
    To         = "johnDoe@Contoso.com";
    Subject    = "Signed Actionable Message";
    Body       = $messageBody;
    BodyAsHtml = $true;
}

Send-MailMessage @parameters