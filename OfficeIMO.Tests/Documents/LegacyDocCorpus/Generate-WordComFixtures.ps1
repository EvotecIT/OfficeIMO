<# 
.SYNOPSIS
Generates small legacy Word 97-2003 .doc fixtures through Microsoft Word COM.

.DESCRIPTION
This script is test-support tooling for the legacy DOC corpus. It intentionally
uses Word automation only outside production code so OfficeIMO.Word stays
dependency-free at runtime.
#>
[CmdletBinding()]
param(
    [string] $OutputDirectory = $PSScriptRoot,

    [ValidateSet('SimpleParagraphs', 'CharacterFormatting', 'ParagraphFormatting')]
    [string[]] $Scenario = @('SimpleParagraphs', 'CharacterFormatting', 'ParagraphFormatting'),

    [switch] $Force
)

$ErrorActionPreference = 'Stop'

$WdDoNotSaveChanges = 0
$WdFormatDocument = 0
$WdStyleNormal = -1
$WdStyleHeading1 = -2
$WdAlignParagraphLeft = 0
$WdAlignParagraphCenter = 1
$WdAlignParagraphRight = 2
$WdAlignParagraphJustify = 3
$WdUnderlineNone = 0
$WdUnderlineSingle = 1
$WdColorAutomatic = -16777216
$WdColorRed = 255
$WdColorBlue = 16711680

function Assert-WindowsPlatform {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'Word COM fixture generation requires Windows with Microsoft Word installed.'
    }
}

function Release-ComObject {
    param(
        [AllowNull()]
        [object] $ComObject
    )

    if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function New-WordApplication {
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        throw 'Microsoft Word COM automation is not available. Install Microsoft Word or run this generator on a workstation that has it.'
    }

    $word.Visible = $false
    $word.DisplayAlerts = 0
    return $word
}

function Reset-SelectionFormat {
    param(
        [object] $Selection
    )

    $Selection.Style = $WdStyleNormal
    $Selection.Font.Bold = 0
    $Selection.Font.Italic = 0
    $Selection.Font.Underline = $WdUnderlineNone
    $Selection.Font.Size = 11
    $Selection.Font.Name = 'Calibri'
    $Selection.Font.Color = $WdColorAutomatic
    $Selection.ParagraphFormat.Alignment = $WdAlignParagraphLeft
    $Selection.ParagraphFormat.SpaceBefore = 0
    $Selection.ParagraphFormat.SpaceAfter = 0
    $Selection.ParagraphFormat.LineSpacing = 12
    $Selection.ParagraphFormat.LeftIndent = 0
    $Selection.ParagraphFormat.RightIndent = 0
    $Selection.ParagraphFormat.FirstLineIndent = 0
}

function Add-WordParagraph {
    param(
        [object] $Selection,
        [string] $Text,
        [scriptblock] $BeforeText
    )

    Reset-SelectionFormat -Selection $Selection
    if ($BeforeText) {
        & $BeforeText $Selection
    }

    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
    Reset-SelectionFormat -Selection $Selection
}

function Save-WordDocument {
    param(
        [object] $Document,
        [string] $Path
    )

    if (Test-Path -LiteralPath $Path) {
        if (-not $Force) {
            Write-Host "Skipping existing fixture: $Path"
            return $false
        }

        Remove-Item -LiteralPath $Path -Force
    }

    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }

    try {
        $Document.SaveAs2([ref] $Path, [ref] $WdFormatDocument)
    } catch {
        $Document.SaveAs([ref] $Path, [ref] $WdFormatDocument)
    }

    Write-Host "Generated fixture: $Path"
    return $true
}

function New-WordComFixture {
    param(
        [object] $Word,
        [string] $Path,
        [scriptblock] $Populate
    )

    if ((Test-Path -LiteralPath $Path) -and -not $Force) {
        Write-Host "Skipping existing fixture: $Path"
        return
    }

    $document = $null
    try {
        $document = $Word.Documents.Add()
        $selection = $Word.Selection
        Reset-SelectionFormat -Selection $selection
        & $Populate $document $selection
        [void](Save-WordDocument -Document $document -Path $Path)
    } finally {
        if ($null -ne $document) {
            $document.Close([ref] $WdDoNotSaveChanges)
        }

        Release-ComObject $document
    }
}

function Add-SimpleParagraphsFixture {
    param(
        [object] $Word,
        [string] $Path
    )

    New-WordComFixture -Word $Word -Path $Path -Populate {
        param($Document, $Selection)

        Add-WordParagraph -Selection $Selection -Text 'First COM paragraph'
        Add-WordParagraph -Selection $Selection -Text 'Second COM paragraph'
    }
}

function Add-CharacterFormattingFixture {
    param(
        [object] $Word,
        [string] $Path
    )

    New-WordComFixture -Word $Word -Path $Path -Populate {
        param($Document, $Selection)

        Add-WordParagraph -Selection $Selection -Text 'COM Heading One' -BeforeText {
            param($Selection)
            $Selection.Style = $WdStyleHeading1
        }

        Add-WordParagraph -Selection $Selection -Text 'bold text' -BeforeText {
            param($Selection)
            $Selection.Font.Bold = 1
        }

        Add-WordParagraph -Selection $Selection -Text 'italic text' -BeforeText {
            param($Selection)
            $Selection.Font.Italic = 1
        }

        Add-WordParagraph -Selection $Selection -Text 'underlined red courier text' -BeforeText {
            param($Selection)
            $Selection.Font.Underline = $WdUnderlineSingle
            $Selection.Font.Color = $WdColorRed
            $Selection.Font.Name = 'Courier New'
            $Selection.Font.Size = 14
        }
    }
}

function Add-ParagraphFormattingFixture {
    param(
        [object] $Word,
        [string] $Path
    )

    New-WordComFixture -Word $Word -Path $Path -Populate {
        param($Document, $Selection)

        Add-WordParagraph -Selection $Selection -Text 'left paragraph'

        Add-WordParagraph -Selection $Selection -Text 'center paragraph' -BeforeText {
            param($Selection)
            $Selection.ParagraphFormat.Alignment = $WdAlignParagraphCenter
        }

        Add-WordParagraph -Selection $Selection -Text 'right paragraph' -BeforeText {
            param($Selection)
            $Selection.ParagraphFormat.Alignment = $WdAlignParagraphRight
        }

        Add-WordParagraph -Selection $Selection -Text 'justified paragraph with spacing and indentation' -BeforeText {
            param($Selection)
            $Selection.ParagraphFormat.Alignment = $WdAlignParagraphJustify
            $Selection.ParagraphFormat.SpaceBefore = 12
            $Selection.ParagraphFormat.SpaceAfter = 6
            $Selection.ParagraphFormat.LineSpacing = 18
            $Selection.ParagraphFormat.LeftIndent = 36
            $Selection.ParagraphFormat.RightIndent = 18
            $Selection.ParagraphFormat.FirstLineIndent = 12
        }
    }
}

Assert-WindowsPlatform

$resolvedOutputDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null

$word = $null
try {
    $word = New-WordApplication

    foreach ($name in $Scenario) {
        switch ($name) {
            'SimpleParagraphs' {
                Add-SimpleParagraphsFixture -Word $word -Path (Join-Path $resolvedOutputDirectory 'ComSimpleParagraphs.doc')
                break
            }
            'CharacterFormatting' {
                Add-CharacterFormattingFixture -Word $word -Path (Join-Path $resolvedOutputDirectory 'ComCharacterFormatting.doc')
                break
            }
            'ParagraphFormatting' {
                Add-ParagraphFormattingFixture -Word $word -Path (Join-Path $resolvedOutputDirectory 'ComParagraphFormatting.doc')
                break
            }
        }
    }
} finally {
    if ($null -ne $word) {
        $word.Quit([ref] $WdDoNotSaveChanges)
    }

    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
