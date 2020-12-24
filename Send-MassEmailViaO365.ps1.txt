# CSV mailer via Office365 SMTP
# Egor Emeliyanov, MIT license

$emailServer    = 'smtp.office365.com'
$emailPort      = 587
$delaySeconds   = 10

function Run-Retriable {

    Param ([ScriptBlock] $command)

    $delays = @(0, 1, 1, 2, 3, 10)
    $tries = 0    

    while ($tries -lt $delays.Count) {
            
        Start-Sleep -Seconds $delays[$tries]
    
        try {            
            & $command
            Return            
        } catch { 
            $tries++
            
            if ($tries -eq $delays.Count) {
                throw
            }
            
            Write-Output "Caught $($_.Exception.Message) retrying..."            
        }    
    }
}

Add-Type -AssemblyName System.Windows.Forms

# prompting for the CSV file
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop');
    Filter = 'CSV files (*.csv)|*.csv';
    Title = "Please select the CSV file with recipient list"
}

$null = $FileBrowser.ShowDialog()

if ($FileBrowser.FileName) {

    $recipients = Import-Csv -Path $FileBrowser.FileName
    Write-Output "Imported $($FileBrowser.FileName)"

} else {

    Write-Output "No input file provided"
    Exit(2)

}

# prompting for the right field in CSV
$recipientsKeys = ($recipients | Get-Member -MemberType NoteProperty).Name

$recipientsKeysSelected = $recipientsKeys | Out-GridView -Title "Please select the email field" -OutputMode Single

if ($recipientsKeysSelected) {

    Write-Output "Selected email field: $recipientsKeysSelected"

} else {

    Write-Output "No recipient key selected"
    Exit(2)
}

# prompting for html file with email body
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop');
    Filter = 'HTML files|*.htm;*.html;*.txt';
    Title = "Please specify the file with HTML body"
}

$null = $FileBrowser.ShowDialog()

if ($FileBrowser.FileName) {

    $HTMLbody = Get-Content -Path $FileBrowser.FileName -Raw 
    Write-Output "Imported HTML $($FileBrowser.FileName)"

} else {

    Write-Output "No input HTML file provided"
    Exit(2)

}

$credentials = Get-Credential -Message "Please enter your O365 credentials. The username will also be used as FROM address."

if (-not $credentials) {
    Write-Output "No credentials supplied"
    Exit(2)
}

$emailFrom = $credentials.UserName

$emailSubject = Read-Host -Prompt "Please enter the subject of email (%variables% will be expanded)"

Write-Output "Total $($recipients.count) entries in the input file"

$recipients = @($recipients | Where-Object {$_.$recipientsKeysSelected -like '*@*.*'})
Write-Output "$($recipients.count) look like valid emails"

 $recipients | Where-Object {$_} | ForEach-Object {

    $currentRecord = $_
    $HTMLbodyCopy = $HTMLbody
    $emailSubjectCopy = $emailSubject
    $emailTo = $currentRecord.$recipientsKeysSelected

    # expanding %variables% in subject and body
    $recipientsKeys | Where-Object {$_} | ForEach-Object {

        $HTMLbodyCopy = $HTMLbodyCopy -replace "%$_%", $currentRecord.$_
        $emailSubjectCopy = $emailSubjectCopy -replace "%$_%", $currentRecord.$_

    }

    $result = Run-Retriable -Command {
        Send-MailMessage -To $emailTo -From $emailFrom -Subject $emailSubjectCopy -Body $HTMLbodyCopy -SmtpServer $emailServer -Port $emailPort -Credential $credentials -BodyAsHtml -UseSsl
    }
    
    Write-Output "Sent email from $emailFrom to $emailTo subject $emailSubjectCopy"
    
    Start-Sleep -Seconds $delaySeconds
}

Write-Output "Normal completion"











