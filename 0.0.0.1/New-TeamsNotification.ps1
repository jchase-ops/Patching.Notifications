# .ExternalHelp $PSScriptRoot\New-TeamsNotification-help.xml
function New-TeamsNotification {

    [CmdletBinding(DefaultParameterSetName = 'Standard')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Standard')]
        [Parameter(Position = 0, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 0, ParameterSetName = 'MultiAction')]
        [Parameter(Position = 0, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Channel,

        [Parameter(Position = 1, ParameterSetName = 'Standard')]
        [Parameter(Position = 1, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 2, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Summary = 'PS Teams Notification',

        [Parameter(Position = 2, ParameterSetName = 'Standard')]
        [Parameter(Position = 2, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 2, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Title,

        [Parameter(Position = 3, ParameterSetName = 'Standard')]
        [Parameter(Position = 3, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 3, ParameterSetName = 'MultiAction')]
        [ValidatePattern('^[a-fA-F0-9]+$')]
        [System.String]
        $ThemeColor,

        [Parameter(Position = 4, ParameterSetName = 'Standard')]
        [Parameter(Position = 4, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 4, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ActivityTitle,

        [Parameter(Position = 5, ParameterSetName = 'Standard')]
        [Parameter(Position = 5, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 5, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ActivitySubtitle,

        [Parameter(Position = 6, ParameterSetName = 'Standard')]
        [Parameter(Position = 6, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 6, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ActivityText,

        [Parameter(Position = 7, ParameterSetName = 'Standard')]
        [Parameter(Position = 7, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $FactTitle,

        [Parameter(Position = 8, ParameterSetName = 'Standard')]
        [Parameter(Position = 8, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $Facts,

        [Parameter(Mandatory, Position = 7, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 7, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $FactSectionList,

        [Parameter(Position = 9, ParameterSetName = 'Standard')]
        [Parameter(Position = 9, ParameterSetName = 'MultiFact')]
        [Parameter(Position = 9, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $FactFormatHash,

        [Parameter(Position = 10, ParameterSetName = 'Standard')]
        [Parameter(Position = 10, ParameterSetName = 'MultiFact')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ActionTitle,

        [Parameter(Position = 11, ParameterSetName = 'Standard')]
        [Parameter(Position = 11, ParameterSetName = 'MultiFact')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $ActionUri,

        [Parameter(Position = 10, ParameterSetName = 'MultiFact')]
        [Parameter(Mandatory, Position = 10, ParameterSetName = 'MultiAction')]
        [ValidateNotNullOrEmpty()]
        [System.Collections.Hashtable[]]
        $ActionSectionList,

        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MultiFact')]
        [Parameter(ParameterSetName = 'MultiAction')]
        [Parameter(ParameterSetName = 'Demo')]
        [Switch]
        $Quiet,

        [Parameter(ParameterSetName = 'Standard')]
        [Parameter(ParameterSetName = 'MultiFact')]
        [Parameter(ParameterSetName = 'MultiAction')]
        [Switch]
        $Test,

        [Parameter(Mandatory, ParameterSetName = 'Demo')]
        [Switch]
        $Demo
    )

    $windowVisible = if ($(Get-Process -Id $([System.Diagnostics.Process]::GetCurrentProcess().Id)).MainWindowHandle -eq 0) { $false } else { $true }

    $formatScriptBlock = [System.Management.Automation.ScriptBlock] {
        Param ($FactList, $FormatHash)
        ForEach ($f in $FactList) {
            $formatKey = $FormatHash.Status | Where-Object { $_ -eq $($f.value.Split(':') | Select-Object -First 1) }
            $formatColor = ($FormatHash | Where-Object { $_.Values -eq $formatKey }).Color
            $f.value = "<font style=`"color:${formatColor}`"><b>$($f.value)</b></font>"
        }
        $FactList
    }

    if (!($Demo)) {
        if (!($Channel)) {
            if ($null -eq $script:Config.Teams.Channel) {
                $script:Config.Teams.Channel = Read-Host -Prompt 'Teams Channel'
            }
        }
        else {
            $script:Config.Teams.Channel = $Channel
        }

        if (!($ThemeColor)) {
            if ($null -eq $script:Config.Teams.ThemeColor) {
                $script:Config.Teams.ThemeColor = Read-Host -Prompt 'ThemeColor (in hex)'
            }
        }
        else {
            $script:Config.Teams.ThemeColor = $ThemeColor
        }

        $script:Config.Teams.Summary = $Summary

        $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100

        if (!($Title)) {
            if ($null -eq $script:Config.Teams.Title) {
                $script:Config.Teams.Title = Read-Host -Prompt 'Title'
            }
        }
        else {
            $script:Config.Teams.Title = $Title
        }

        if (!($ActivityTitle)) {
            if ($null -eq $script:Config.Teams.ActivityTitle) {
                $script:Config.Teams.ActivityTitle = Read-Host -Prompt 'ActivityTitle'
            }
        }
        else {
            $script:Config.Teams.ActivityTitle = $ActivityTitle
        }

        if (!($ActivitySubtitle)) {
            if ($null -eq $script:Config.Teams.ActivitySubtitle) {
                $script:Config.Teams.ActivitySubtitle = Read-Host -Prompt 'ActivitySubtitle'
            }
        }
        else {
            $script:Config.Teams.ActivitySubtitle = $ActivitySubtitle
        }

        if (!($ActivityText)) {
            if ($null -eq $script:Config.Teams.ActivityText) {
                $script:Config.Teams.ActivityText = Read-Host -Prompt 'ActivityText'
            }
        }
        else {
            $script:Config.Teams.ActivityText = $ActivityText
        }

        $body = [PSCustomObject]@{
            '@type'         = 'MessageCard'
            '@context'      = 'http://schema.org/extensions'
            themeColor      = $script:Config.Teams.ThemeColor
            summary         = $script:Config.Teams.Summary
            title           = $script:Config.Teams.Title
            sections        = [System.Collections.Generic.List[System.Object]]::New()
            potentialAction = [System.Collections.Generic.List[System.Object]]::New()
        }
        $activitySection = @{
            activityTitle    = $script:Config.Teams.ActivityTitle
            activitySubtitle = $script:Config.Teams.ActivitySubtitle
            activityText     = $script:Config.Teams.ActivityText
        }
        $body.sections.Add($activitySection)

        if ($PSCmdlet.ParameterSetName -ne 'MultiFact') {
            $factSection = @{
                title = $script:Config.Teams.FactTitle
                facts = $null
            }
            if ($FactFormatHash) {
                $factSection.facts = $formatScriptBlock.Invoke($Facts, $FactFormatHash)
            }
            else {
                $factSection.facts = $Facts
            }
            $body.sections.Add($factSection)
        }
        else {
            ForEach ($fs in $FactSectionList) {
                $factSection = @{
                    title = $fs.Title
                    facts = $null
                }
                if ($null -ne $fs.FactFormatHash) {
                    $factSection.facts = $formatScriptBlock.Invoke($fs.Facts, $fs.FactFormatHash)
                }
                else {
                    $factSection.facts = $fs.Facts
                }
                $body.sections.Add($factSection)
            }
        }

        if ($PSCmdlet.ParameterSetName -ne 'MultiAction') {
            if ($ActionTitle) {
                $actionSection = @{
                    '@context' = 'http://schema.org'
                    '@type'    = 'OpenUri'
                    name       = $ActionTitle
                    targets    = $null 
                }
                if ($ActionUri) { $actionSection.targets = @(@{ os = 'default'; uri = $ActionUri }) }
                $body.potentialAction.Add($actionSection)
            }
        }
        else {
            ForEach ($as in $ActionSectionList) {
                $actionSection = @{
                    '@context' = 'http://schema.org'
                    '@type'    = 'OpenUri'
                    name       = $as.ActionTitle
                    targets    = $null
                }
                if ($null -ne $as.ActionUri) { $actionSection.targets = @(@{ os = 'default'; uri = $as.ActionUri }) }
                $body.potentialAction.Add($actionSection)
            }
        }

        $r = Invoke-RestMethod -Uri $script:Config.Teams.Channel -Method POST -Body $($body | ConvertTo-Json -Depth 100) -ContentType 'application/json'
        if ($windowVisible -and !($Quiet)) {
            Write-Host 'Teams Notification - ' -NoNewline
            if ($r -eq 1) {
                Write-Host 'Sent' -ForegroundColor Green
            }
            else {
                Write-Host 'Failed' -ForegroundColor Red
                Write-Host $r -ForegroundColor Yellow
            }
        }
    }
    else {
        if (!($Channel)) {
            if ($null -eq $script:Demo.Teams.Channel) {
                $script:Demo.Teams.Channel = Read-Host -Prompt 'Teams Channel'
            }
        }
        else {
            $script:Demo.Teams.Channel = $Channel
        }

        $body = [PSCustomObject]@{
            '@type'         = 'MessageCard'
            '@context'      = 'http://schema.org/extensions'
            themeColor      = $(Get-Random -InputObject $script:Demo.Teams.ThemeColor -Count 6) -Join ''
            summary         = $script:Demo.Teams.Summary
            title           = $script:Demo.Teams.Title
            sections        = [System.Collections.Generic.List[System.Object]]::New()
            potentialAction = [System.Collections.Generic.List[System.Object]]::New()
        }
        $activitySection = @{
            activityTitle    = $script:Demo.Teams.ActivityTitle
            activitySubtitle = $script:Demo.Teams.ActivitySubtitle
            activityText     = $script:Demo.Teams.ActivityText
        }
        $body.sections.Add($activitySection)
        ForEach ($key in $script:HTMLColor.Keys) {
            $factSection = @{
                title = $key
                facts = [System.Collections.Generic.List[System.Collections.Hashtable]]::New()
            }
            ForEach ($color in $script:HTMLColor.$key.Keys) {
                $factSection.facts.Add(@{ name = $color; value = "<font style=`"color:${color}`"><b>${color}</b></font>" })
            }
            $body.sections.Add($factSection)
        }
        $actionSection = @{
            '@context' = 'http://schema.org'
            '@type'    = 'OpenUri'
            name       = $script:Demo.Teams.ActionTitle
            targets    = @(@{ os = 'default'; uri = $script:Demo.Teams.ActionUri })
        }
        $body.potentialAction.Add($actionSection)

        $r = Invoke-RestMethod -Uri $script:Demo.Teams.Channel -Method POST -Body $($body | ConvertTo-Json -Depth 100) -ContentType 'application/json'
        if ($windowVisible -and !($Quiet)) {
            Write-Host 'Teams Notification - ' -NoNewline
            if ($r -eq 1) {
                Write-Host 'Sent' -ForegroundColor Green
            }
            else {
                Write-Host 'Failed' -ForegroundColor Red
                Write-Host $r -ForegroundColor Yellow
            }
        }
    }
}
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUzm+dE1TBtbpn0FLyIJ5+6rFM
# DGKgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUgwyHW3jHXvv4fTTm1Kxbo7pM
# S0gwDQYJKoZIhvcNAQEBBQAEggEAEA2H2/IAIWPD7tAlmitjovj/7L9HSLzllexo
# oCCpJ66kMSeN6Iiz/NR826TltNUQymIPzc7W8OeAbwdRVdYGQudPXhKD7/gP+POq
# PLdBMS3rjOIgCR5RZ/0KTBNFzeYr7VS7wZSrt7OCfJ8FYexa7r+aPSSdaVJJLNz3
# M48xFF3zi6G36fDGZSmJcqiBjK/wlC94PNcvYo/MfUM723/SYVnwPMEkyvwIOP4s
# g4ziK58RlQZDEIh2my9FxhLxgWZYwCSisEllqeIdFY3QIsS50mt6Kh9F9sZP6O2x
# 0HRWQ8nLfrMhFoD4uJXzRL79HhO74P6PSzC8IIjz70zlZDefZA==
# SIG # End signature block
