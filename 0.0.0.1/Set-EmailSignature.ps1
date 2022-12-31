# .ExternalHelp $PSScriptRoot\Set-EmailSignature-help.xml
function Set-EmailSignature {

    [CmdletBinding(DefaultParameterSetName = 'Standard')]

    Param (

        [Parameter(Position = 0, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String[]]
        $ADProperties = @('Department', 'GivenName', 'MobilePhone', 'OfficePhone', 'PostalCode', 'State', 'StreetAddress', 'Surname', 'Title'),

        [Parameter(Position = 1, ParameterSetName = 'Standard')]
        [Parameter(Position = 1, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Domain,

        [Parameter(Position = 2, ParameterSetName = 'Standard')]
        [Parameter(Position = 2, ParameterSetName = 'Demo')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $EmailSuffix,

        [Parameter(Position = 3, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $AreaCode,

        [Parameter(Position = 4, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $CompanyName,

        [Parameter(Position = 5, ParameterSetName = 'Standard')]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $City,

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

    if (!($Domain)) {
        if ($windowVisible -and !($Quiet)) {
            $Domain = Read-Host -Prompt 'Enter Email Domain'
        }
    }

    if (!($EmailSuffix)) {
        if ($windowVisible -and !($Quiet)) {
            $EmailSuffix = Read-Host -Prompt 'Enter Email Suffix'
        }
    }

    $emailAddress = "${env:USERNAME}@$($Domain.ToLower()).${EmailSuffix}"
    $script:Config.Email.From = $emailAddress
    $script:Demo.Email.From = $emailAddress

    if (!($Demo)) {
        if (!($AreaCode)) {
            if ($windowVisible -and !($Quiet)) {
                $AreaCode = Read-Host -Prompt 'Enter Area Code'
            }
        }

        if (!($CompanyName)) {
            if ($windowVisible -and !($Quiet)) {
                $CompanyName = Read-Host -Prompt 'Enter Company Name'
            }
        }

        if (!($City)) {
            if ($windowVisible -and !($Quiet)) {
                $City = Read-Host -Prompt 'Enter City'
            }
        }

        $currentUser = Get-ADUser -Identity $env:USERNAME -Properties $ADProperties

        if ($currentUser.OfficePhone.Length -eq 7) {
            $officePhone = "${AreaCode}-$($currentUser.OfficePhone.Substring(0,3))-$($currentUser.OfficePhone.Substring(3,4))"
        }
        elseif ($currentUser.OfficePhone.Length -eq 10) {
            $officePhone = "${AreaCode}-$($currentUser.OfficePhone.Substring(3,3))-$($currentUser.OfficePhone.Substring(6,4))"
        }
        else {
            $officePhone = $currentUser.OfficePhone -Replace '[()-]\s?', ''
            $officePhone = "${AreaCode}-$($officePhone.Substring(3,3))-$($officePhone.Substring(6,4))"
        }
    }
    else {
        $currentUser = [PSCustomObject]@{
            Department    = $script:Demo.CurrentUser.Department
            GivenName     = $script:Demo.CurrentUser.GivenName
            MobilePhone   = $script:Demo.CurrentUser.MobilePhone
            OfficePhone   = $script:Demo.CurrentUser.OfficePhone
            PostalCode    = $script:Demo.CurrentUser.PostalCode
            State         = $script:Demo.CurrentUser.State
            StreetAddress = $script:Demo.CurrentUser.StreetAddress
            Surname       = $script:Demo.CurrentUser.Surname
            Title         = $script:Demo.CurrentUser.Title
        }

        $officePhone = $currentUser.OfficePhone
        $CompanyName = $script:Demo.CompanyName
        $City = $script:Demo.City
    }

    $sign = New-Object -TypeName System.Text.StringBuilder
    [void]$sign.Append('<br>Respectfully,<br><br>')
    [void]$sign.Append("<b><u><span style='font-size:14.0pt;font-family:`"Century Gothic`",sans-serif;color:black'>$($currentUser.GivenName) $($currentUser.Surname)</span></u></b><br>")
    [void]$sign.Append("<b><span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:gray'>$($currentUser.Title)</span></b><br>")
    [void]$sign.Append("<b><span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:gray'>$($currentUser.Department)</span></b><br>")
    [void]$sign.Append("<b><span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:gray'>${CompanyName}</span></b><br>")
    [void]$sign.Append("<span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:gray'>$($currentUser.StreetAddress) | ${City}, $($currentUser.State) | $($currentUser.PostalCode)</span><br>")
    [void]$sign.Append("<span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:gray'>Office: ${officePhone} | Cell: $($currentUser.MobilePhone)</span><br>")
    [void]$sign.Append("<u><span style='font-size:10.0pt;font-family:`"Century Gothic`",sans-serif;color:#0070C0'><a href='mailto:${emailAddress}'><span style='color=#0070C0'>${emailAddress}</span></a></span></u>")
    
    if ($Demo -or $Test) {
        $sign.ToString()
    }
    else {
        $script:Config.Signature = $sign.ToString()
        $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
    }
}
