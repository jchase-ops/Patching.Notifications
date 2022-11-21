# Patching.Notifications Module

#region Classes
################################################################################
#                                                                              #
#                                 CLASSES                                      #
#                                                                              #
################################################################################
# . "$PSScriptRoot\$(Split-Path -Path $(Split-Path -Path $PSScriptRoot -Parent) -Leaf).Classes.ps1"
#endregion

#region Variables
################################################################################
#                                                                              #
#                               VARIABLES                                      #
#                                                                              #
################################################################################
try {
    $script:Config = Import-Clixml -Path "$PSScriptRoot\config.xml" -ErrorAction Stop
}
catch {
    $script:Config = [ordered]@{
        Email     = [ordered]@{
            Bcc                        = $null
            BodyAsHtml                 = $null
            Cc                         = $null
            Css                        = $null
            DeliveryNotificationOption = $null
            From                       = $null
            SmtpServer                 = $null
            Subject                    = $null
            To                         = $null
            UseSsl                     = $null
        }
        Teams     = [ordered]@{
            ActionUri        = $null
            ActivityTitle    = $null
            ActivitySubtitle = $null
            ActivityText     = $null
            Channel          = $null
            FactFormatHash   = $null
            FactTitle        = $null
            Facts            = $null
            Summary          = $null
            ThemeColor       = $null
            Title            = $null
        }
        Body      = $null
        Signature = $null
    }
    $script:Config | Export-Clixml -Path "$PSScriptRoot\config.xml" -Depth 100
}
finally {
    $script:Demo = [ordered]@{
        CurrentUser = [ordered]@{
            Department    = 'IT Operations'
            GivenName     = 'Joshua'
            MobilePhone   = '602-867-5309'
            OfficePhone   = '123-456-7890'
            PostalCode    = '90210'
            State         = 'AZ'
            StreetAddress = '2138 S Noche de Paz'
            Surname       = 'Chase'
            Title         = 'Script Kitty'
        }
        CompanyName = 'ACME Corporation'
        City        = 'Frankfurt'
        Email       = [ordered]@{
            Attachments                = "$PSScriptRoot\Demo\Demo_Attachment_1.txt", "$PSScriptRoot\Demo\Demo_Attachment_2.txt"
            Bcc                        = $null
            Body                       = Get-Content -Path "$PSScriptRoot\Demo\demo.html"
            BodyAsHtml                 = $true
            Cc                         = $null
            Css                        = Get-Content -Path "$PSScriptRoot\Demo\demo.css"
            DeliveryNotificationOption = 'None'
            Encoding                   = 'Default'
            From                       = $null
            Port                       = '25'
            Priority                   = 'High'
            SmtpServer                 = $null
            Subject                    = 'Demo Email Notification'
            To                         = $null
            UseSsl                     = $true
        }
        Teams       = [ordered]@{
            ActionUri        = 'https://www.lmgtfy.com'
            ActivityTitle    = "$($script:HTMLColor.Keys.Count) Color Groups!"
            ActivitySubtitle = "$($script:HTMLColor.Values.Keys.Count) Color Names!"
            ActivityText     = "$(($script:HTMLColor.Values.Values | Sort-Object -Unique).Count) Colors!"
            Channel          = $null
            Summary          = 'PS Teams Notification'
            ThemeColor       = '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A', 'B', 'C', 'D', 'E', 'F'
            Title            = "Demo Teams Notification"
        }
    }
    $script:HTMLColor = [ordered]@{
        Blue   = [ordered]@{
            Aqua            = '00FFFF'
            Aquamarine      = '7FFFD4'
            Blue            = '0000FF'
            CadetBlue       = '5F9EA0'
            CornflowerBlue  = '6495ED'
            Cyan            = '00FFFF'
            DarkBlue        = '00008B'
            DarkTurquoise   = '00CED1'
            DeepSkyBlue     = '00BFFF'
            DodgerBlue      = '1E90FF'
            LightBlue       = 'ADD8E6'
            LightCyan       = 'E0FFFF'
            LightSkyBlue    = '87CEFA'
            LightSteelBlue  = 'B0C4DE'
            MediumBlue      = '0000CD'
            MediumSlateBlue = '7B68EE'
            MediumTurquoise = '48D1CC'
            MidnightBlue    = '191970'
            Navy            = '000080'
            PaleTurquoise   = 'AFEEEE'
            PowderBlue      = 'B0E0E6'
            RoyalBlue       = '4169E1'
            SkyBlue         = '87CEEB'
            SteelBlue       = '4682B4'
            Turquoise       = '40E0D0'
        }
        Brown  = [ordered]@{
            BlanchedAlmond = 'FFEBCD'
            Bisque         = 'FFE4C4'
            Brown          = 'A52A2A'
            BurlyWood      = 'DEB887'
            Chocolate      = 'D2691E'
            Cornsilk       = 'FFF8DC'
            DarkGoldenrod  = 'B8860B'
            Goldenrod      = 'DAA520'
            Maroon         = '800000'
            NavajoWhite    = 'FFDEAD'
            Peru           = 'CD853F'
            RosyBrown      = 'BC8F8F'
            SaddleBrown    = '8B4513'
            SandyBrown     = 'F4A460'
            Sienna         = 'A0522D'
            Tan            = 'D2B48C'
            Wheat          = 'F5DEB3'
        }
        Gray   = [ordered]@{
            Black          = '000000'
            DarkGray       = 'A9A9A9'
            DarkSlateGray  = '2F4F4F'
            DimGray        = '696969'
            Gainsboro      = 'DCDCDC'
            Gray           = '808080'
            LightGray      = 'D3D3D3'
            LightSlateGray = '708090'
            Silver         = 'C0C0C0'
        }
        Green  = [ordered]@{
            Chartreuse        = '7FFF00'
            DarkCyan          = '008B8B'
            DarkGreen         = '006400'
            DarkOliveGreen    = '556B2F'
            DarkSeaGreen      = '8FBC8B'
            ForestGreen       = '228B22'
            Green             = '008000'
            GreenYellow       = 'ADFF2F'
            LawnGreen         = '7CFC00'
            LightGreen        = '90EE90'
            LightSeaGreen     = '20B2AA'
            Lime              = '00FF00'
            LimeGreen         = '32CD32'
            MediumAquamarine  = '66CDAA'
            MediumSeaGreen    = '3CB371'
            MediumSpringGreen = '00FA9A'
            Olive             = '808000'
            OliveDrab         = '6B8E23'
            PaleGreen         = '98FB98'
            SeaGreen          = '2E8B57'
            SpringGreen       = '00FF7F'
            Teal              = '008080'
            YellowGreen       = '9ACD32'
        }
        Orange = [ordered]@{
            Coral       = 'FF7F50'
            DarkOrange  = 'FF8C00'
            LightSalmon = 'FFA07A'
            Orange      = 'FFA500'
            OrangeRed   = 'FF4500'
            Tomato      = 'FF6347'
        }
        Pink   = [ordered]@{
            DeepPink        = 'FF1493'
            HotPink         = 'FF69B4'
            LightPink       = 'FFB6C1'
            MediumVioletRed = 'C71585'
            PaleVioletRed   = 'DB7093'
            Pink            = 'FFC0CB'
        }
        Purple = [ordered]@{
            BlueViolet      = '8A2BE2'
            DarkOrchid      = '9932CC'
            DarkMagenta     = '8B008B'
            DarkSlateBlue   = '483D8B'
            DarkViolet      = '9400D3'
            Fuchsia         = 'FF00FF'
            Indigo          = '4B0082'
            Lavender        = 'E6E6FA'
            Magenta         = 'FF00FF'
            MediumOrchid    = 'BA55D3'
            MediumPurple    = '9370DB'
            MediumSlateBlue = '7B68EE'
            Orchid          = 'DA70D6'
            Plum            = 'DDA0DD'
            Purple          = '800080'
            RebeccaPurple   = '663399'
            SlateBlue       = '6A5ACD'
            Thistle         = 'D8BFD8'
            Violet          = 'EE82EE'
        }
        Red    = [ordered]@{
            Crimson     = 'DC143C'
            DarkRed     = '8B0000'
            DarkSalmon  = 'E9967A'
            FireBrick   = 'B22222'
            IndianRed   = 'CD5C5C'
            LightCoral  = 'F08080'
            LightSalmon = 'FFA07A'
            Red         = 'FF0000'
            Salmon      = 'FA8072'
        }
        White  = [ordered]@{
            AliceBlue     = 'F0F8FF'
            AntiqueWhite  = 'FAEBD7'
            Azure         = 'F0FFFF'
            Beige         = 'F5F5DC'
            FloralWhite   = 'FFFAF0'
            GhostWhite    = 'F8F8FF'
            HoneyDew      = 'F0FFF0'
            Ivory         = 'FFFFF0'
            LavenderBlush = 'FFF0F5'
            Linen         = 'FAF0E6'
            MintCream     = 'F5FFFA'
            MistyRose     = 'FFE4E1'
            OldLace       = 'FDF5E6'
            SeaShell      = 'FFF5EE'
            Snow          = 'FFFAFA'
            White         = 'FFFFFF'
            WhiteSmoke    = 'F5F5F5'
        }
        Yellow = [ordered]@{
            DarkKhaki            = 'BDB76B'
            Gold                 = 'FFD700'
            Khaki                = 'F0E68C'
            LemonChiffon         = 'FFFACD'
            LightGoldenrodYellow = 'FAFAD2'
            LightYellow          = 'FFFFE0'
            Mocassin             = 'FFE4B5'
            PaleGoldenrod        = 'EEE8AA'
            PapayaWhip           = 'FFEFD5'
            PeachPuff            = 'FFDAB9'
            Yellow               = 'FFFF00'
        }
    }
}
#endregion

#region DotSourcedScripts
################################################################################
#                                                                              #
#                           DOT-SOURCED SCRIPTS                                #
#                                                                              #
################################################################################
. "$PSScriptRoot\New-EmailNotification.ps1"
. "$PSScriptRoot\New-TeamsNotification.ps1"
. "$PSScriptRoot\Set-EmailSignature.ps1"
#endregion

#region ModuleMembers
################################################################################
#                                                                              #
#                              MODULE MEMBERS                                  #
#                                                                              #
################################################################################
Export-ModuleMember -Function New-EmailNotification
Export-ModuleMember -Function New-TeamsNotification
Export-ModuleMember -Function Set-EmailSignature
#endregion
# SIG # Begin signature block
# MIIFYQYJKoZIhvcNAQcCoIIFUjCCBU4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUVGSRYvuUixoohoUYu0PxnkhG
# uqKgggMAMIIC/DCCAeSgAwIBAgIQFnL4oVNG56NIRjNfzwNXejANBgkqhkiG9w0B
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
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUtUVHQ6FPV5uNPLC+acsnWdcm
# kAkwDQYJKoZIhvcNAQEBBQAEggEAws+skiPg53Ug70ZcYFLP3ws6VwfdtmYtYggS
# 7u0kDhEM0I5FXBvcgRvNhc7wGRcSmfxfwFCOmrRlreNlDreKKD+bYWdMxiwvMzOC
# zpLUetiCJPWj+pnREqjzytAjaWK7fIsA/4G67Mpe1litI5bIJqF60rzF3+N0+y/9
# gQs6IhykdYUVPcaCV2TQ1WPgZp4k8zkART6xTcr5WgLoSAFFdTtxgeiJgrK/GhxN
# /INolA4ycrKFS3k1LxNWKDonKna3XmgGl6WcP50Yc45o/gkeRN154ZOvtR2YG27t
# lcjrzJ6OYo5u5HicTM+iurj+AgFFWoZbz4a2Eo0mbmzrNzk1rA==
# SIG # End signature block
