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
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUEKnYQY+bhl5+dg9qGF9WPN1r
# 91WgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
# AQUFADAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MDAeFw0yMTEyMDIwNDU5MTJaFw0y
# MjEyMDIwNTE5MTJaMBYxFDASBgNVBAMMC0NlcnQtMDM0NTYwMIIBIjANBgkqhkiG
# 9w0BAQEFAAOCAQ8AMIIBCgKCAQEA8daSAcUBI0Xx8sMMlSpsCV+24lY46RsxX8iC
# bB7ZM19b/GBjwMo0TCb28ssbZ/P8liNJICrSbyIkQDrIrjqtAdyAPdPAYHONTHad
# 0fuOQQT5MkO5HAxUYLz/6H/xq92lKQFxz5Wgzw+3KOyignY8V8ZZ379z/WqQbNCV
# +29zb9YWOK7eXQ9x8s4+SOizqUE3zkOuijf86I9vZmzMYhsxE7if0R0UlQsLlvTA
# kH/m4IjHem8rl/kC+O71lU7l9475XrUUR3Fxebqh9YoCEZh2eE81TLQcnvK8zgqP
# F+X4INdNPD6zO4T1Nbz0Ccev7mj37+pk/eL5R5aV+NJgqAzhvQIDAQABo0YwRDAO
# BgNVHQ8BAf8EBAMCBaAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFFNN
# e4x6JSqbcnTR354fVSEgQ0VYMA0GCSqGSIb3DQEBBQUAA4IBAQBXfA8VgaMD2c/v
# Sv8gnS/LWri51BBqcUFE9JYMxEIzlEt2ZfJsG+INaQqzBoyCDx/oMQH7wdFRvDjQ
# QsXpNTo7wH7WytFe9KJrOz2uGG0EnIYHK0dTFIMVOcM9VsWWPG40EAzD//55xX/d
# pBL+L4SSTujbR3ptni8Agu5GiRhTpxwl1L/HLC2QYYMoUKiAxL1p61+cHRj6wMzl
# jxnrMIcBhKioaXnwWdKPCN66Jk8IYdqr8afcRYiwtDi+8Hk2/9nB9HwPox3Dtf8H
# jH0O2/8NiJTeOBFSfrWPM9r4j4NWR8IuLwsqHUfXJEQa9SOxhHvxaNMR/Fhq1GVj
# qUClZiXiMYIByzCCAccCAQEwKjAWMRQwEgYDVQQDDAtDZXJ0LTAzNDU2MAIQFnL4
# oVNG56NIRjNfzwNXejAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUdJpy2wD51kA+SyLtPlvIA3MU
# IAQwDQYJKoZIhvcNAQEBBQAEggEAE3i2hWxlQ5A/ZaqFhOVpmouDWacLpZXn1r1u
# 85XL7NiciT9vKQelxj7HbkR48mmtEG2iPTjKxjq3a/amyrQT6qSHRujP0a9JOTPt
# tbdA7eVDcC3EAWSReIynHBmz46y8/kSGMTXnFSl3JKkf44JLKFHePLt8tH8M+bbh
# 0VtOfzhPnz9ZI6nTnI7IUuVdf8kCv9L1zU7G8/AbEtx/8sbJ88SNrUTntbp9PqWM
# z6KoAKAIqXW35cz6lS7rGQ/nnsV1OgQC0lXuyta3B6wWW5D0b/bsfIfgrPW5ZmRK
# r9mG4PBOvD2T6qtyU6GnOWKyd3HB5OCkL+FLGgA2375xqF9X6w==
# SIG # End signature block
