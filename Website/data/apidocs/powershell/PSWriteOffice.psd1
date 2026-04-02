@{
    AliasesToExport        = @('ConvertFrom-WordHtml', 'ConvertFrom-WordMarkdown', 'ConvertTo-MarkdownHtml', 'ConvertTo-WordHtml', 'ConvertTo-WordMarkdown', 'ExcelAutoFilter', 'ExcelAutoFilterClear', 'ExcelAutoFit', 'ExcelCell', 'ExcelChart', 'ExcelColumn', 'ExcelComment', 'ExcelCommentRemove', 'ExcelConditionalColorScale', 'ExcelConditionalDataBar', 'ExcelConditionalIconSet', 'ExcelConditionalRule', 'ExcelFormula', 'ExcelFreeze', 'ExcelGridlines', 'ExcelHeaderFooter', 'ExcelHyperlink', 'ExcelHyperlinkHost', 'ExcelHyperlinkSmart', 'ExcelImage', 'ExcelImageFromUrl', 'ExcelInternalLinks', 'ExcelInternalLinksByHeader', 'ExcelMargins', 'ExcelNamedRange', 'ExcelOrientation', 'ExcelPageSetup', 'ExcelPivotTable', 'ExcelPivotTables', 'ExcelProtect', 'ExcelRow', 'ExcelSheet', 'ExcelSheetVisibility', 'ExcelSort', 'ExcelSparkline', 'ExcelTable', 'ExcelTableOfContents', 'ExcelUnprotect', 'ExcelUrlLinks', 'ExcelUrlLinksByHeader', 'ExcelValidationCustomFormula', 'ExcelValidationDate', 'ExcelValidationDecimal', 'ExcelValidationList', 'ExcelValidationTextLength', 'ExcelValidationTime', 'ExcelValidationWholeNumber', 'MarkdownCallout', 'MarkdownCode', 'MarkdownDefinitionList', 'MarkdownDetails', 'MarkdownFrontMatter', 'MarkdownHeading', 'MarkdownHorizontalRule', 'MarkdownHr', 'MarkdownImage', 'MarkdownList', 'MarkdownParagraph', 'MarkdownQuote', 'MarkdownTable', 'MarkdownTableOfContents', 'MarkdownTaskList', 'MarkdownToc', 'PptBackground', 'PptBullets', 'PptChart', 'PptImage', 'PptLayoutBox', 'PptLayoutPlaceholderBounds', 'PptLayoutPlaceholderMargins', 'PptLayoutPlaceholders', 'PptLayoutPlaceholderTextStyle', 'PptNotes', 'PptPlaceholderText', 'PptShape', 'PptSlide', 'PptSlideLayout', 'PptSlideSize', 'PptTable', 'PptTextBox', 'PptTheme', 'PptThemeColor', 'PptThemeFonts', 'PptThemeName', 'PptTitle', 'PptTransition', 'Replace-OfficePowerPointText', 'WordBold', 'WordBookmark', 'WordCheckBox', 'WordCheckBoxes', 'WordComboBox', 'WordComboBoxes', 'WordContentControl', 'WordContentControls', 'WordDatePicker', 'WordDatePickers', 'WordDropDownList', 'WordDropDownLists', 'WordField', 'WordFooter', 'WordHeader', 'WordHyperlink', 'WordImage', 'WordItalic', 'WordList', 'WordListItem', 'WordPageNumber', 'WordParagraph', 'WordPictureControl', 'WordPictureControls', 'WordRepeatingSection', 'WordRepeatingSections', 'WordSection', 'WordTable', 'WordTableCondition', 'WordTableOfContent', 'WordTableOfContents', 'WordText', 'WordWatermark')
    Author                 = 'Przemyslaw Klys'
    CmdletsToExport        = @('Add-OfficeExcelAutoFilter', 'Add-OfficeExcelChart', 'Add-OfficeExcelComment', 'Add-OfficeExcelConditionalColorScale', 'Add-OfficeExcelConditionalDataBar', 'Add-OfficeExcelConditionalIconSet', 'Add-OfficeExcelConditionalRule', 'Add-OfficeExcelImage', 'Add-OfficeExcelImageFromUrl', 'Add-OfficeExcelPivotTable', 'Add-OfficeExcelSheet', 'Add-OfficeExcelSparkline', 'Add-OfficeExcelTable', 'Add-OfficeExcelTableOfContents', 'Add-OfficeExcelValidationCustomFormula', 'Add-OfficeExcelValidationDate', 'Add-OfficeExcelValidationDecimal', 'Add-OfficeExcelValidationList', 'Add-OfficeExcelValidationTextLength', 'Add-OfficeExcelValidationTime', 'Add-OfficeExcelValidationWholeNumber', 'Add-OfficeMarkdownCallout', 'Add-OfficeMarkdownCode', 'Add-OfficeMarkdownDefinitionList', 'Add-OfficeMarkdownDetails', 'Add-OfficeMarkdownFrontMatter', 'Add-OfficeMarkdownHeading', 'Add-OfficeMarkdownHorizontalRule', 'Add-OfficeMarkdownImage', 'Add-OfficeMarkdownList', 'Add-OfficeMarkdownParagraph', 'Add-OfficeMarkdownQuote', 'Add-OfficeMarkdownTable', 'Add-OfficeMarkdownTableOfContents', 'Add-OfficeMarkdownTaskList', 'Add-OfficePowerPointBullets', 'Add-OfficePowerPointChart', 'Add-OfficePowerPointImage', 'Add-OfficePowerPointSection', 'Add-OfficePowerPointShape', 'Add-OfficePowerPointSlide', 'Add-OfficePowerPointTable', 'Add-OfficePowerPointTextBox', 'Add-OfficeWordBookmark', 'Add-OfficeWordCheckBox', 'Add-OfficeWordComboBox', 'Add-OfficeWordContentControl', 'Add-OfficeWordDatePicker', 'Add-OfficeWordDropDownList', 'Add-OfficeWordField', 'Add-OfficeWordFooter', 'Add-OfficeWordHeader', 'Add-OfficeWordHyperlink', 'Add-OfficeWordImage', 'Add-OfficeWordList', 'Add-OfficeWordListItem', 'Add-OfficeWordPageNumber', 'Add-OfficeWordParagraph', 'Add-OfficeWordPictureControl', 'Add-OfficeWordRepeatingSection', 'Add-OfficeWordSection', 'Add-OfficeWordTable', 'Add-OfficeWordTableCondition', 'Add-OfficeWordTableOfContent', 'Add-OfficeWordText', 'Add-OfficeWordWatermark', 'Clear-OfficeExcelAutoFilter', 'Close-OfficeExcel', 'Close-OfficeWord', 'ConvertFrom-OfficeWordHtml', 'ConvertFrom-OfficeWordMarkdown', 'ConvertTo-OfficeCsv', 'ConvertTo-OfficeMarkdown', 'ConvertTo-OfficeMarkdownHtml', 'ConvertTo-OfficeWordHtml', 'ConvertTo-OfficeWordMarkdown', 'Copy-OfficePowerPointSlide', 'Find-OfficeWord', 'Get-OfficeCsv', 'Get-OfficeCsvData', 'Get-OfficeExcel', 'Get-OfficeExcelData', 'Get-OfficeExcelNamedRange', 'Get-OfficeExcelPivotTable', 'Get-OfficeExcelRange', 'Get-OfficeExcelTable', 'Get-OfficeExcelUsedRange', 'Get-OfficeMarkdown', 'Get-OfficePowerPoint', 'Get-OfficePowerPointLayout', 'Get-OfficePowerPointLayoutBox', 'Get-OfficePowerPointLayoutPlaceholder', 'Get-OfficePowerPointNotes', 'Get-OfficePowerPointPlaceholder', 'Get-OfficePowerPointSection', 'Get-OfficePowerPointShape', 'Get-OfficePowerPointSlide', 'Get-OfficePowerPointSlideSummary', 'Get-OfficePowerPointTheme', 'Get-OfficeWord', 'Get-OfficeWordBookmark', 'Get-OfficeWordCheckBox', 'Get-OfficeWordComboBox', 'Get-OfficeWordContentControl', 'Get-OfficeWordDatePicker', 'Get-OfficeWordDocumentProperty', 'Get-OfficeWordDropDownList', 'Get-OfficeWordField', 'Get-OfficeWordHyperlink', 'Get-OfficeWordParagraph', 'Get-OfficeWordPictureControl', 'Get-OfficeWordRepeatingSection', 'Get-OfficeWordRun', 'Get-OfficeWordSection', 'Get-OfficeWordTable', 'Get-OfficeWordTableOfContent', 'Import-OfficePowerPointSlide', 'Invoke-OfficeExcelAutoFit', 'Invoke-OfficeExcelSort', 'Invoke-OfficeWordMailMerge', 'New-OfficeExcel', 'New-OfficeMarkdown', 'New-OfficePowerPoint', 'New-OfficeWord', 'Protect-OfficeExcelSheet', 'Protect-OfficeWordDocument', 'Remove-OfficeExcelComment', 'Remove-OfficePowerPointSlide', 'Remove-OfficeWordTableOfContent', 'Rename-OfficePowerPointSection', 'Save-OfficeExcel', 'Save-OfficePowerPoint', 'Save-OfficeWord', 'Set-OfficeExcelCell', 'Set-OfficeExcelChartDataLabels', 'Set-OfficeExcelChartLegend', 'Set-OfficeExcelChartStyle', 'Set-OfficeExcelColumn', 'Set-OfficeExcelFormula', 'Set-OfficeExcelFreeze', 'Set-OfficeExcelGridlines', 'Set-OfficeExcelHeaderFooter', 'Set-OfficeExcelHostHyperlink', 'Set-OfficeExcelHyperlink', 'Set-OfficeExcelInternalLinks', 'Set-OfficeExcelInternalLinksByHeader', 'Set-OfficeExcelMargins', 'Set-OfficeExcelNamedRange', 'Set-OfficeExcelOrientation', 'Set-OfficeExcelPageSetup', 'Set-OfficeExcelRow', 'Set-OfficeExcelSheetVisibility', 'Set-OfficeExcelSmartHyperlink', 'Set-OfficeExcelUrlLinks', 'Set-OfficeExcelUrlLinksByHeader', 'Set-OfficePowerPointBackground', 'Set-OfficePowerPointLayoutPlaceholderBounds', 'Set-OfficePowerPointLayoutPlaceholderTextMargins', 'Set-OfficePowerPointLayoutPlaceholderTextStyle', 'Set-OfficePowerPointNotes', 'Set-OfficePowerPointPlaceholderText', 'Set-OfficePowerPointSlideLayout', 'Set-OfficePowerPointSlideSize', 'Set-OfficePowerPointSlideTitle', 'Set-OfficePowerPointSlideTransition', 'Set-OfficePowerPointThemeColor', 'Set-OfficePowerPointThemeFonts', 'Set-OfficePowerPointThemeName', 'Set-OfficeWordBackground', 'Set-OfficeWordDocumentProperty', 'Set-OfficeWordTableOfContent', 'Unprotect-OfficeExcelSheet', 'Update-OfficePowerPointText', 'Update-OfficeWordFields', 'Update-OfficeWordTableOfContent')
    CompanyName            = 'Evotec'
    CompatiblePSEditions   = @('Desktop', 'Core')
    Copyright              = '(c) 2011 - 2026 Przemyslaw Klys @ Evotec. All rights reserved.'
    Description            = 'PowerShell module to create and read Microsoft Word, Excel, PowerPoint (experimental), Markdown, and CSV documents without Microsoft Office installed. Powered by OfficeIMO.*.'
    DotNetFrameworkVersion = '4.7.2'
    FunctionsToExport      = @()
    GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
    ModuleVersion          = '0.3.0'
    PowerShellVersion      = '5.1'
    PrivateData            = @{
        PSData = @{
            LicenseUri                 = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'
            ProjectUri                 = 'https://github.com/EvotecIT/PSWriteOffice'
            Tags                       = @('officeimo', 'word', 'excel', 'powerpoint', 'markdown', 'csv', 'docx', 'xlsx', 'pptx', 'openxml', 'windows', 'linux', 'macos')
            RequireLicenseAcceptance   = $false
            ExternalModuleDependencies = @()
        }
    }
    RootModule             = 'PSWriteOffice.psm1'
    RequiredModules        = @()
    ScriptsToProcess       = @()
}

# SIG # Begin signature block
# MIIt6AYJKoZIhvcNAQcCoIIt2TCCLdUCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUwXeZ1KmMpopL
# 26Ceh4vKfu9QknP/xl1PFLtIOh5T6qCCJuUwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggWQMIIDeKADAgECAhAFmxtXno4hMuI5B72nd3VcMA0GCSqG
# SIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0xMzA4MDExMjAwMDBaFw0zODAxMTUxMjAwMDBaMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL/mkHNo3rvkXUo8MCIw
# aTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3EMB/zG6Q4FutWxpdtHauyefLK
# EdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKyunWZanMylNEQRBAu34LzB4Tm
# dDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsFxl7sWxq868nPzaw0QF+xembu
# d8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU15zHL2pNe3I6PgNq2kZhAkHnD
# eMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJBMtfbBHMqbpEBfCFM1LyuGwN1
# XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObURWBf3JFxGj2T3wWmIdph2PVld
# QnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6nj3cAORFJYm2mkQZK37AlLTS
# YW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxBYKqxYxhElRp2Yn72gLD76GSm
# M9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5SUUd0viastkF13nqsX40/ybzT
# QRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+xq4aLT8LWRV+dIPyhHsXAj6Kx
# fgommfXkaS+YHS312amyHeUbAgMBAAGjQjBAMA8GA1UdEwEB/wQFMAMBAf8wDgYD
# VR0PAQH/BAQDAgGGMB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwPTzANBgkq
# hkiG9w0BAQwFAAOCAgEAu2HZfalsvhfEkRvDoaIAjeNkaA9Wz3eucPn9mkqZucl4
# XAwMX+TmFClWCzZJXURj4K2clhhmGyMNPXnpbWvWVPjSPMFDQK4dUPVS/JA7u5iZ
# aWvHwaeoaKQn3J35J64whbn2Z006Po9ZOSJTROvIXQPK7VB6fWIhCoDIc2bRoAVg
# X+iltKevqPdtNZx8WorWojiZ83iL9E3SIAveBO6Mm0eBcg3AFDLvMFkuruBx8lbk
# apdvklBtlo1oepqyNhR6BvIkuQkRUNcIsbiJeoQjYUIp5aPNoiBB19GcZNnqJqGL
# FNdMGbJQQXE9P01wI4YMStyB0swylIQNCAmXHE/A7msgdDDS4Dk0EIUhFQEI6FUy
# 3nFJ2SgXUE3mvk3RdazQyvtBuEOlqtPDBURPLDab4vriRbgjU2wGb2dVf0a1TD9u
# KFp5JtKkqGKX0h7i7UqLvBv9R0oN32dmfrJbQdA75PQ79ARj6e/CVABRoIoqyc54
# zNXqhwQYs86vSYiv85KZtrPmYQ/ShQDnUBrkG5WdGaG5nLGbsQAe79APT0JsyQq8
# 7kP6OnGlyE0mpTX9iV28hWIdMtKgK1TtmlfB2/oQzxm3i0objwG2J5VT6LaJbVu8
# aNQj6ItRolb58KaAoNYes7wPD1N1KarqE3fk3oyBIa0HEEcRrYc9B9F1vM/zZn4w
# ggawMIIEmKADAgECAhAIrUCyYNKcTJ9ezam9k67ZMA0GCSqGSIb3DQEBDAUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yMTA0MjkwMDAwMDBaFw0zNjA0MjgyMzU5NTlaMGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDVtC9C0CiteLdd1TlZG7GIQvUz
# jOs9gZdwxbvEhSYwn6SOaNhc9es0JAfhS0/TeEP0F9ce2vnS1WcaUk8OoVf8iJnB
# kcyBAz5NcCRks43iCH00fUyAVxJrQ5qZ8sU7H/Lvy0daE6ZMswEgJfMQ04uy+wjw
# iuCdCcBlp/qYgEk1hz1RGeiQIXhFLqGfLOEYwhrMxe6TSXBCMo/7xuoc82VokaJN
# TIIRSFJo3hC9FFdd6BgTZcV/sk+FLEikVoQ11vkunKoAFdE3/hoGlMJ8yOobMubK
# wvSnowMOdKWvObarYBLj6Na59zHh3K3kGKDYwSNHR7OhD26jq22YBoMbt2pnLdK9
# RBqSEIGPsDsJ18ebMlrC/2pgVItJwZPt4bRc4G/rJvmM1bL5OBDm6s6R9b7T+2+T
# YTRcvJNFKIM2KmYoX7BzzosmJQayg9Rc9hUZTO1i4F4z8ujo7AqnsAMrkbI2eb73
# rQgedaZlzLvjSFDzd5Ea/ttQokbIYViY9XwCFjyDKK05huzUtw1T0PhH5nUwjeww
# k3YUpltLXXRhTT8SkXbev1jLchApQfDVxW0mdmgRQRNYmtwmKwH0iU1Z23jPgUo+
# QEdfyYFQc4UQIyFZYIpkVMHMIRroOBl8ZhzNeDhFMJlP/2NPTLuqDQhTQXxYPUez
# +rbsjDIJAsxsPAxWEQIDAQABo4IBWTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAd
# BgNVHQ4EFgQUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHwYDVR0jBBgwFoAU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJo
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# bDAcBgNVHSAEFTATMAcGBWeBDAEDMAgGBmeBDAEEATANBgkqhkiG9w0BAQwFAAOC
# AgEAOiNEPY0Idu6PvDqZ01bgAhql+Eg08yy25nRm95RysQDKr2wwJxMSnpBEn0v9
# nqN8JtU3vDpdSG2V1T9J9Ce7FoFFUP2cvbaF4HZ+N3HLIvdaqpDP9ZNq4+sg0dVQ
# eYiaiorBtr2hSBh+3NiAGhEZGM1hmYFW9snjdufE5BtfQ/g+lP92OT2e1JnPSt0o
# 618moZVYSNUa/tcnP/2Q0XaG3RywYFzzDaju4ImhvTnhOE7abrs2nfvlIVNaw8rp
# avGiPttDuDPITzgUkpn13c5UbdldAhQfQDN8A+KVssIhdXNSy0bYxDQcoqVLjc1v
# djcshT8azibpGL6QB7BDf5WIIIJw8MzK7/0pNVwfiThV9zeKiwmhywvpMRr/Lhlc
# OXHhvpynCgbWJme3kuZOX956rEnPLqR0kq3bPKSchh/jwVYbKyP/j7XqiHtwa+ag
# uv06P0WmxOgWkVKLQcBIhEuWTatEQOON8BUozu3xGFYHKi8QxAwIZDwzj64ojDzL
# j4gLDb879M4ee47vtevLt/B3E+bnKD+sEq6lLyJsQfmCXBVmzGwOysWGw/YmMwwH
# S6DTBwJqakAwSEs0qFEgu60bhQjiWQ1tygVQK+pKHJ6l/aCnHwZ05/LWUpD9r4VI
# IflXO7ScA+2GRfS0YW6/aOImYIbqyK+p/pQd52MbOoZWeE4wgga0MIIEnKADAgEC
# AhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcw
# MDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1l
# U3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0ZodLRRF51NrY0NlLWZ
# loMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi6wuim5bap+0lgloM
# 2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNgxVBdJkf77S2uPoCj
# 7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQ
# Sku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJm/s80FiocSk1VYLZ
# lDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+
# 8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1ZlAeSpQl92QOMeRx
# ykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9MmeOreGPRdtBx3yG
# OP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7Yk/ehb//Wx+5kMqI
# MRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm
# 1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6X5uAiynM7Bu2ayBj
# UwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU729T
# SunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4c
# D08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUF
# BwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEG
# CCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRU
# cnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAX
# MAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBABfO+xaA
# HP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQ
# M2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt
# 6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0F8HABBgr0UdqirZ7
# bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmS
# Nq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKfZxAvBAKqMVuqte69
# M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzEwlvzZiiyfTPjLbnF
# RsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmM
# Thi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOXgpIUsWTjd6xpR6oa
# Qf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx
# 9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3
# /BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIG7TCCBNWgAwIBAgIQCoDvGEuN
# 8QWC0cR2p5V0aDANBgkqhkiG9w0BAQsFADBpMQswCQYDVQQGEwJVUzEXMBUGA1UE
# ChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQg
# VGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMB4XDTI1MDYwNDAw
# MDAwMFoXDTM2MDkwMzIzNTk1OVowYzELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRp
# Z2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJEaWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBU
# aW1lc3RhbXAgUmVzcG9uZGVyIDIwMjUgMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBANBGrC0Sxp7Q6q5gVrMrV7pvUf+GcAoB38o3zBlCMGMyqJnfFNZx
# +wvA69HFTBdwbHwBSOeLpvPnZ8ZN+vo8dE2/pPvOx/Vj8TchTySA2R4QKpVD7dvN
# Zh6wW2R6kSu9RJt/4QhguSssp3qome7MrxVyfQO9sMx6ZAWjFDYOzDi8SOhPUWlL
# nh00Cll8pjrUcCV3K3E0zz09ldQ//nBZZREr4h/GI6Dxb2UoyrN0ijtUDVHRXdmn
# cOOMA3CoB/iUSROUINDT98oksouTMYFOnHoRh6+86Ltc5zjPKHW5KqCvpSduSwhw
# UmotuQhcg9tw2YD3w6ySSSu+3qU8DD+nigNJFmt6LAHvH3KSuNLoZLc1Hf2JNMVL
# 4Q1OpbybpMe46YceNA0LfNsnqcnpJeItK/DhKbPxTTuGoX7wJNdoRORVbPR1VVnD
# uSeHVZlc4seAO+6d2sC26/PQPdP51ho1zBp+xUIZkpSFA8vWdoUoHLWnqWU3dCCy
# FG1roSrgHjSHlq8xymLnjCbSLZ49kPmk8iyyizNDIXj//cOgrY7rlRyTlaCCfw7a
# SUROwnu7zER6EaJ+AliL7ojTdS5PWPsWeupWs7NpChUk555K096V1hE0yZIXe+gi
# AwW00aHzrDchIc2bQhpp0IoKRR7YufAkprxMiXAJQ1XCmnCfgPf8+3mnAgMBAAGj
# ggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0GA1UdDgQWBBTkO/zyMe39/dfzkXFjGVBD
# z2GM6DAfBgNVHSMEGDAWgBTvb1NK6eQGfHrK4pBW9i/USezLTjAOBgNVHQ8BAf8E
# BAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGF
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXQYIKwYBBQUH
# MAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRH
# NFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBW
# MFSgUqBQhk5odHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVk
# RzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkw
# FzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQBlKq3x
# HCcEua5gQezRCESeY0ByIfjk9iJP2zWLpQq1b4URGnwWBdEZD9gBq9fNaNmFj6Eh
# 8/YmRDfxT7C0k8FUFqNh+tshgb4O6Lgjg8K8elC4+oWCqnU/ML9lFfim8/9yJmZS
# e2F8AQ/UdKFOtj7YMTmqPO9mzskgiC3QYIUP2S3HQvHG1FDu+WUqW4daIqToXFE/
# JQ/EABgfZXLWU0ziTN6R3ygQBHMUBaB5bdrPbF6MRYs03h4obEMnxYOX8VBRKe1u
# NnzQVTeLni2nHkX/QqvXnNb+YkDFkxUGtMTaiLR9wjxUxu2hECZpqyU1d0IbX6Wq
# 8/gVutDojBIFeRlqAcuEVT0cKsb+zJNEsuEB7O7/cuvTQasnM9AWcIQfVjnzrvwi
# CZ85EE8LUkqRhoS3Y50OHgaY7T/lwd6UArb+BOVAkg2oOvol/DJgddJ35XTxfUlQ
# +8Hggt8l2Yv7roancJIFcbojBcxlRcGG0LIhp6GvReQGgMgYxQbV1S3CrWqZzBt1
# R9xJgKf47CdxVRd/ndUlQ05oxYy2zRWVFjF7mcr4C34Mj3ocCVccAvlKV9jEnstr
# niLvUxxVZE/rptb7IRE2lskKPIJgbaP5t2nGj/ULLi49xTcBZU8atufk+EMF/cWu
# iC7POGT75qaL6vdCvHlshtjdNXOCIUjsarfNZzCCB18wggVHoAMCAQICEAfCUnQo
# FKLWq/4k6hfl3S4wDQYJKoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNV
# BAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0
# IENvZGUgU2lnbmluZyBSU0E0MDk2IFNIQTM4NCAyMDIxIENBMTAeFw0yMzA0MTYw
# MDAwMDBaFw0yNjA3MDYyMzU5NTlaMGcxCzAJBgNVBAYTAlBMMRIwEAYDVQQHDAlN
# aWtvxYLDs3cxITAfBgNVBAoMGFByemVteXPFgmF3IEvFgnlzIEVWT1RFQzEhMB8G
# A1UEAwwYUHJ6ZW15c8WCYXcgS8WCeXMgRVZPVEVDMIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEAlJoHlzELSGimkpCr2wLfBhWSdcsDh/EsMZU7rODHMq1p
# lTq0QVUUAPAKRfRWnqG8JpGcb5MUExSxypvvJJ8KJhFLJXGvAqkjiNGMBC7+RME1
# RIdAvw2nob8aOrZJjTxff0j9Sgt3NJdbzvjO73TVRikCEK4cauxBtInswWTgIrpD
# XRlV0WDi5+O1d6i+T8Bv6LtmpSf74nyA2nfNahW/kJFIdNiaNuEjI1nSg8rXazF4
# tNt+QjeEa1vvII30Sfnyio4DCJm7nHgrIvSL9Wuum1HPWpwHpjm0+JheVP8kAYAL
# gKN/o1QfMIlHfO5FEDtMyQhfL6tmK1Ts/DiZjF/IICLBBFGdwmSg9IVXN3Zu3Fkg
# MPPxTcxjT5QGiMc11/ang9BIGgi0ZCLQN7d3kFviAF8kv/WZ56RVKA70BmyvkOP2
# z9Im/fFy30KcVRkbtHAldDYO+wyJERfiMkdT3MFQKvjs1VN7ynqNub/657Ylwpgs
# YluKB2DtvHkkP3iAHJ4ovt7igzWayNeT+1cQ65FCHOhbYkrzocHNwM2PrxH4r1JB
# SkasL0kq+Hwq65JO89kHu9mcJcNhA0VR8stH1FRjvUDLoehN0cJyS/eoqdGpXJoS
# gARqCKkltOZ13QlG5F5oTwk0+Z2kA7mdVJAF22T0oSo2z8M3Vz9m/CPZ0PPVUoEC
# AwEAAaOCAgMwggH/MB8GA1UdIwQYMBaAFGg34Ou2O/hfEYb7/mF7CIhl9E5CMB0G
# A1UdDgQWBBR68WolWbgyccRJNeWy6DLhSOdt9zAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwgbUGA1UdHwSBrTCBqjBToFGgT4ZNaHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0
# MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BRoE+GTWh0dHA6Ly9jcmw0LmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNBNDA5NlNIQTM4NDIw
# MjFDQTEuY3JsMD4GA1UdIAQ3MDUwMwYGZ4EMAQQBMCkwJwYIKwYBBQUHAgEWG2h0
# dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCBlAYIKwYBBQUHAQEEgYcwgYQwJAYI
# KwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBcBggrBgEFBQcwAoZQ
# aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29k
# ZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcnQwCQYDVR0TBAIwADANBgkq
# hkiG9w0BAQsFAAOCAgEAtxHh11D4aXt9Stgy+Nx34eqpLwR8kdUZQ/ZVSJJXEQke
# dGR86FrOhAZUxcqIb5KXJVQrkXUFt97Uur7SjzrnKQw7+MLAPus5CWCPHx6Lluk6
# mtVuO2Eq3OQDkoSHCffjaTWyjRood3aEpXIqNplCgl+SP2a8yQZEKSdJGIWv6VEk
# 9gmxNya6CX9r0FhlIiPidy3YjzR5oTtZfs2kJEsb9HFQxEzH0BmSikVREmehYOtW
# 9HY70EseddDHW8bSjI70t2bQMrap0B5NYqT/kYPjOZRR60pFJZ6Rmvn957kIcQ2+
# zfRPIVFXr8QC06xYn4PM4bJVUR+fw3/wsZTClwu6Kd9PwMkLDkMR1tbjcd7RtQzI
# Is6cAWrK8YesGu4mgPi6dO6tSPdni4a2G7cN8QtrzSBnTHTe53e+sjCI3WJwJ+69
# /MMLWidymA9EE5e+xAfLv+XArN0oWXQ3coOCuzaCZfIhB626raKABzjC4iaYi9ov
# WJ/JEDAev0OkTDtyFDy7snAfaOgzYsEB3+ibeaFuz9PZOTccQRJpLMcDW5mbzUOu
# WZ93sVACqhvsd9RIM+SGeFP4z80WpRJRCKUtK4K1YPEfKRDoXfeZhM6eVhEShcl4
# Xupwem0mB7/HJSwFdIjJLt9PK6X4zIkJKktyy831CeTh6rSikDTC8c/c9fOVArMx
# ggZZMIIGVQIBATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEAfCUnQoFKLWq/4k6hfl3S4wDQYJYIZI
# AWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0B
# CQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAv
# BgkqhkiG9w0BCQQxIgQgmMgFoVo/2U1S4vbASsEUoK/mb+8vphSfRRteNbztc+Ew
# DQYJKoZIhvcNAQEBBQAEggIAKwbVZD66tgpGh3sV+iOC3njwFQ/K+7EGoRMM6IKV
# x3PWhUUmYlPQtAy+a9zrYi3dh+toEmskWIWJZQvUVrLYSXn1VATdNo4lxz5K4WNa
# Cu8+PJySWU5I/7GKjIpy/sx8JinZwT/UDz39ZIK1buZOXFUu6RLwsTKg2yi+IUj+
# jlG8sBMGbQ9sjJNQiLyFE8wZdnmqikC7bvMBmCmzHmuOPRH8kpoF6iUqslOGT3WJ
# lsPumiGKgEHd6RgRqs1zKZsFVIf7+YkMkMKYfP2Jel3gNeQlbWEqapdb79wDAL37
# mhjtfa5sd1q8wb+84zretN4IeGnRNipzJ2LYBpHVrxDTTAqv+tQZ0vyRSVklRd+g
# df1h35m+mRIljriDfPJZ4q+y3uP05zZKVVCJTodnRRJ3NGj6/xEnXy5hOb+fd9oN
# pyXO2DJznc91siP2AQ3gan5UEumGtT3TlSkcJANPWkCTpv3R0koJLhQLs0OmS0zH
# 7mlmOKxrp3h35ivBAEyjTul3aGqtrmgJuUYV+9eBkyxirx3dUl0KSYfLavmubbfC
# TuFwuR7aDq7jOquY4pACD7a5VIFCJlBxJHxW/ibmR4RZIyUjPsPOviQpONmoKg+U
# LqdO72kKYxhnq0iWGLITxI7D95DfwWwgXdapTb90YJ0XLH2lQLXtydtQ5tFGh300
# 3aOhggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRy
# dXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExAhAK
# gO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZIhvcNAQkDMQsG
# CSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYwNDAxMTc1MzU1WjAvBgkqhkiG
# 9w0BCQQxIgQg6flAZK+l37AtXdHmTrulsf5PzuH1mp5R+K/Z4b7jPPUwDQYJKoZI
# hvcNAQEBBQAEggIAxcGgAkmyctTUCCSvhi64REwp7/D85WaaRm3A1rc5YcLQ2z1+
# 7h/7NBLKWPqMq1f7K4UW+cPj0LosyVzmyCFMuBlxzuINSknqXqhBMiIkms2JQNq+
# JTu/4Lvg9ueRLBYVvdgzYOj43Jgufgh/TUcvyh5a3JCQ20kTPzZRkdSY7UOHiOOT
# gu2u1V2vJGSJHjRlrFUug4mm5hkTEDitL+z3lcLIvbmsjFqB5BMpgiHgEcXX3e2T
# 586sbKXUF04soVlif6mb6iv2W2bfkUTp4c3OeIVcG8ZnxE9hx6zKtN+2p06qA38R
# swszTnejbvAITpETIRbQU78eVZ200p3L9/VbcWxD+t7GNRpGIFG6n4b0pY3KkB72
# HcsrKnVgWWPi9Shy68osZIU+6gia/95seBKG9EjdUAbWf464/KqpjmHMeFnNvhB0
# UTN2Mqpi8F+/Tmt9537EsuohfP2zcpHsgfCYuPftq4/Bm22cXWF4+dLAcFCck3GC
# Si8D61UvRGi5u71lwQsr5ITEQq0Vn3KWqe9vQUcwrnYgx2liBTexw79+cxtH8WMB
# 97kT5il/S0d/qnHhSE70CrS95rDEEmYRNvcrR0mYQFcbJv1zSpH3Scy1H+Hz/xUF
# RvZVAbcgPFagfPD8OiDwCijE7kkS8l4FeAdudRlHdyQA/28Ix/7sTNW7acs=
# SIG # End signature block
