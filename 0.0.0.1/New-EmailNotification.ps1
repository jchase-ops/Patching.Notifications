# .ExternalHelp $PSScriptRoot\New-EmailNotification-help.xml
function New-EmailNotification {

    [CmdletBinding(DefaultParameterSetName = 'Standard')]

    Param (

        [Parameter(Mandatory, Position = 0, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Subject,

        [Parameter(Mandatory, Position = 1, ParameterSetName = 'Standard')]
        [Parameter(Position = 1, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $To,

        [Parameter(Position = 2, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Attachments,

        [Parameter(Position = 3, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Bcc,

        [Parameter(Mandatory, Position = 4, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Body,

        [Parameter(ParameterSetName = 'Standard')]
        [Switch]
        $BodyAsHtml,

        [Parameter(Position = 5, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $Cc,

        [Parameter(Position = 6, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Css,

        [Parameter(Position = 7, ParameterSetName = 'Standard')]
        [ValidateSet('None', 'OnSuccess', 'OnFailure', 'Delay', 'Never')]
        [System.String]
        $DeliveryNotificationOption = 'None',

        [Parameter(Position = 8, ParameterSetName = 'Standard')]
        [ValidateSet('ASCII', 'BigEndianUnicode', 'Default', 'OEM', 'Unicode', 'UTF7', 'UTF8', 'UTF32')]
        [System.String]
        $Encoding = 'Default',

        [Parameter(Position = 9, ParameterSetName = 'Standard')]
        [Parameter(Position = 9, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $From,

        [Parameter(Position = 10, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Port,

        [Parameter(Position = 11, ParameterSetName = 'Standard')]
        [ValidateSet('Normal', 'High', 'Low')]
        [System.String]
        $Priority,

        [Parameter(Position = 12, ParameterSetName = 'Standard')]
        [Parameter(Position = 12, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $SmtpServer,

        [Parameter(ParameterSetName = 'Standard')]
        [Switch]
        $UseSsl,

        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'Demo')]
        [Switch]
        $Quiet,

        [Parameter(ParameterSetName = 'Standard')]
        [Switch]
        $Test,

        [Parameter(ParameterSetName = 'Demo')]
        [Switch]
        $Demo
    )

    $windowVisible = if ($(Get-Process -Id $([System.Diagnostics.Process]::GetCurrentProcess().Id)).MainWindowHandle -eq 0) { $false } else { $true }

    if (!($Demo)) {
        if (!($From)) {
            if ($null -eq $script:Config.Email.From) {
                if ($null -ne $script:Config.Signature) {
                    $script:Config.Email.From = (($script:Config.Signature.Split('<br>') | Select-String -Pattern "mailto:" -SimpleMatch).Line.Split(':') | Select-Object -Last 1).TrimEnd("'")
                }
                else {
                    if (!($Quiet)) {
                        Set-EmailSignature
                    }
                }
            }
        }
        else {
            $script:Config.Email.From = $From
        }

        if (!($SmtpServer)) {
            if ($null -eq $script:Config.Email.SmtpServer) {
                if ($windowVisible -and !($Quiet)) {
                    $script:Config.Email.SmtpServer = Read-Host -Prompt 'SMTP Server'
                }
                else {
                    exit
                }
            }
        }
        else {
            $script:Config.Email.SmtpServer = $SmtpServer
        }

        $emailParams = [System.Collections.Hashtable]::New()
        $emailParams.Subject = $Subject
        $emailParams.To = $To
        $emailParams.From = $script:Config.Email.From
        $emailParams.SmtpServer = $script:Config.Email.SmtpServer

        if ($Attachments) { $emailParams.Attachments = $Attachments }
        if ($Bcc) { $emailParams.Bcc = $Bcc }
        if ($BodyAsHtml) {
            $emailParams.BodyAsHtml = $true
            $emailBody = New-Object -TypeName System.Text.StringBuilder
            [void]$emailBody.Append('<!DOCTYPE html><html><head><style>')
            if ($Css) {
                [void]$emailBody.Append("$($Css)</style></head><body>$($Body)<br>$($script:Config.Signature)</body></html>")
            }
            else {
                [void]$emailBody.Append("</style></head><body>$($Body)<br>$($script:Config.Signature)</body></html>")
            }
            $emailParams.Body = $emailBody.ToString()
        }
        else {
            $emailParams.Body = $Body
        }
        if ($Cc) { $emailParams.Cc = $true }
        if ($DeliveryNotificationOption) { $emailParams.DeliveryNotificationOption = $DeliveryNotificationOption }
        if ($Encoding) { $emailParams.Encoding = $Encoding }
        if ($Port) { $emailParams.Port = $Port }
        if ($Priority) { $emailParams.Priority = $Priority }
        if ($UseSsl) { $emailParams.UseSsl = $true }

        if ($Test) {
            $emailParams.To = $script:Config.Email.From
            $emailParams.Bcc = $script:Config.Email.From
            $emailParams.Cc = $script:Config.Email.From
        }
    }
    else {
        if (!($From)) {
            if ($null -eq $script:Demo.Email.From) {
                if ($null -ne $script:Config.Signature) {
                    $script:Demo.Email.From = (($script:Config.Signature.Split('<br>') | Select-String -Pattern "mailto:" -SimpleMatch).Line.Split(':') | Select-Object -Last 1).TrimEnd("'")
                }
                else {
                    if (!($Quiet)) {
                        Set-EmailSignature
                    }
                }
            }
        }
        else {
            $script:Demo.Email.From = $From
        }

        if (!($SmtpServer)) {
            if ($null -eq $script:Demo.Email.SmtpServer) {
                if ($windowVisible -and !($Quiet)) {
                    $script:Demo.Email.SmtpServer = Read-Host -Prompt 'SMTP Server'
                }
                else {
                    exit
                }
            }
        }
        else {
            $script:Demo.Email.SmtpServer = $SmtpServer
        }

        $emailBody = New-Object -TypeName System.Text.StringBuilder
        [void]$emailBody.Append("<!DOCTYPE html><html><head><style>$($script:Demo.Email.Css)</style></head><body>$($script:Demo.Email.Body)<br>$($script:Config.Signature)</body></html>")

        $emailParams = [System.Collections.Hashtable]::New()
        $emailParams.Subject = $script:Demo.Email.Subject
        $emailParams.To = $script:Demo.Email.From
        $emailParams.From = $script:Demo.Email.From
        $emailParams.SmtpServer = $script:Demo.Email.SmtpServer
        $emailParams.Attachments = $script:Demo.Email.Attachments
        $emailParams.Bcc = $script:Demo.Email.From
        $emailParams.Body = $emailBody.ToString()
        $emailParams.BodyAsHtml = $script:Demo.Email.BodyAsHtml
        $emailParams.Cc = $script:Demo.Email.From
        $emailParams.DeliveryNotificationOption = $script:Demo.Email.DeliveryNotificationOption
        $emailParams.Encoding = $script:Demo.Email.Encoding
        $emailParams.Port = $script:Demo.Email.Port
        $emailParams.Priority = $script:Demo.Email.Priority
        $emailParams.UseSsl = $script:Demo.Email.UseSsl
    }

    Send-MailMessage @emailParams
    if ($?) {
        if ($windowVisible -and !($Quiet)) {
            Write-Host 'Email Notification - ' -NoNewline
            Write-Host 'Sent' -ForegroundColor Green
        }
        else {
            Write-Host 'Email Notification - ' -NoNewline
            Write-Host 'Failed' -ForegroundColor Red
        }
    }
}
