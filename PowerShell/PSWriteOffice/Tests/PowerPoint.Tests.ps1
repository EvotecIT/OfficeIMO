BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

    function Test-OfficeLoadedMethod {
        param(
            [Parameter(Mandatory)]
            [string] $TypeName,

            [Parameter(Mandatory)]
            [string] $MethodName
        )

        $type = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType($TypeName, $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        if ($null -eq $type) {
            throw "Unable to find loaded type '$TypeName'."
        }

        @($type.GetMethods() | Where-Object Name -eq $MethodName).Count -gt 0
    }
}

Describe 'PowerPoint cmdlets' {
    It 'does not create a NoSave presentation when WhatIf skips creation' {
        $path = Join-Path $TestDrive 'PowerPointNoSaveWhatIf.pptx'

        New-OfficePowerPoint -Path $path -NoSave -WhatIf | Out-Null

        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'creates the parent directory before opening a NoSave presentation' {
        $folder = Join-Path $TestDrive 'missing'
        $path = Join-Path $folder 'PowerPointNoSave.pptx'

        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $presentation | Should -Not -BeNullOrEmpty
            Test-Path -LiteralPath $folder | Should -BeTrue
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation -ErrorAction SilentlyContinue
            }
        }
    }

    It 'closes a presentation through a PowerShell cmdlet' {
        $path = Join-Path $TestDrive 'PowerPointClose.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        Add-OfficePowerPointSlide -Presentation $presentation | Out-Null

        Close-OfficePowerPoint -Presentation $presentation -Save

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'saves to a new path while closing a presentation' {
        $sourcePath = Join-Path $TestDrive 'PowerPointCloseSource.pptx'
        $savedAsPath = Join-Path (Join-Path $TestDrive 'close-save-as') 'PowerPointClose.Copy.pptx'
        $presentation = New-OfficePowerPoint -Path $sourcePath -NoSave
        Add-OfficePowerPointSlide -Presentation $presentation

        Close-OfficePowerPoint -Presentation $presentation -Path $savedAsPath

        Test-Path -LiteralPath $savedAsPath | Should -BeTrue
        $reloaded = Get-OfficePowerPoint -Path $savedAsPath
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'saves an external OfficeIMO presentation without closing it and supports save as' {
        $savedAsPath = Join-Path (Join-Path $TestDrive 'save-as-output') 'PowerPointExternal.Copy.pptx'
        $seedPath = Join-Path $TestDrive 'PowerPointTypeSeed.pptx'
        $seed = New-OfficePowerPoint -Path $seedPath -NoSave
        $presentationType = $seed.GetType()
        Close-OfficePowerPoint -Presentation $seed
        $create = $presentationType.GetMethods() |
            Where-Object { $_.Name -eq 'Create' -and $_.GetParameters().Count -eq 0 } |
            Select-Object -First 1
        $presentation = $create.Invoke($null, @())
        try {
            Add-OfficePowerPointSlide -Presentation $presentation | Out-Null

            $returned = Save-OfficePowerPoint -Presentation $presentation -Path $savedAsPath -PassThru
            [object]::ReferenceEquals($presentation, $returned) | Should -BeTrue
            Test-Path -LiteralPath $savedAsPath | Should -BeTrue

            Add-OfficePowerPointSlide -Presentation $presentation | Out-Null
            Save-OfficePowerPoint -Presentation $presentation
            $presentation.Slides.Count | Should -Be 2
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation -ErrorAction SilentlyContinue
            }
        }

        $reloaded = Get-OfficePowerPoint -Path $savedAsPath
        try {
            $reloaded.Slides.Count | Should -Be 2
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'does not persist edits made after save when closed without save' {
        $path = Join-Path $TestDrive 'PowerPointExplicitClose.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        Add-OfficePowerPointSlide -Presentation $presentation | Out-Null
        Save-OfficePowerPoint -Presentation $presentation
        Add-OfficePowerPointSlide -Presentation $presentation | Out-Null
        Close-OfficePowerPoint -Presentation $presentation

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'rejects external SaveOnDispose presentations before tracking them' {
        $path = Join-Path $TestDrive 'PowerPointExternalAutoSave.pptx'
        $presentationType = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType('OfficeIMO.PowerPoint.PowerPointPresentation', $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        $optionsType = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType('OfficeIMO.PowerPoint.PowerPointCreateOptions', $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        $modeType = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType('OfficeIMO.DocumentPersistenceMode', $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        $options = [Activator]::CreateInstance($optionsType)
        $options.PersistenceMode = [Enum]::Parse($modeType, 'SaveOnDispose')
        $create = $presentationType.GetMethods() | Where-Object {
            $_.Name -eq 'Create' -and
            $_.GetParameters().Count -eq 2 -and
            $_.GetParameters()[0].ParameterType -eq [string]
        } | Select-Object -First 1
        $presentation = $create.Invoke($null, [object[]]@([string] $path, $options))
        try {
            Add-OfficePowerPointSlide -Presentation $presentation | Out-Null

            { Save-OfficePowerPoint -Presentation $presentation -Path $path -ErrorAction Stop } |
                Should -Throw '*DocumentPersistenceMode.Explicit*'
            { Close-OfficePowerPoint -Presentation $presentation -ErrorAction Stop } |
                Should -Throw '*DocumentPersistenceMode.Explicit*'
        } finally {
            $presentation.Dispose()
        }

        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'treats case-distinct encrypted source paths separately on case-sensitive systems' {
        if (-not [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
                [System.Runtime.InteropServices.OSPlatform]::Linux)) {
            return
        }
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.PowerPoint.PowerPointPresentation' -MethodName 'OpenEncrypted')) {
            return
        }

        $sourcePath = Join-Path $TestDrive 'Deck.pptx'
        $savedAsPath = Join-Path $TestDrive 'deck.pptx'
        New-OfficePowerPoint -Path $sourcePath -Password 'secret' {
            PptSlide { PptTitle -Title 'Encrypted source' }
        }

        $presentation = Get-OfficePowerPoint -Path $sourcePath -Password 'secret'
        try {
            Save-OfficePowerPoint -Presentation $presentation -Path $savedAsPath
        } finally {
            Close-OfficePowerPoint -Presentation $presentation -ErrorAction SilentlyContinue
        }

        Test-Path -LiteralPath $sourcePath | Should -BeTrue
        Test-Path -LiteralPath $savedAsPath | Should -BeTrue
        $encryptedReload = Get-OfficePowerPoint -Path $sourcePath -Password 'secret'
        try {
            $encryptedReload.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $encryptedReload
        }
        $reloaded = Get-OfficePowerPoint -Path $savedAsPath
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'saves an encrypted presentation through the canonical path writer' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.PowerPoint.PowerPointPresentation' -MethodName 'OpenEncrypted')) {
            return
        }

        $sourcePath = Join-Path $TestDrive 'EncryptedPowerPointSource.pptx'
        $savedAsPath = Join-Path (Join-Path $TestDrive 'encrypted-output') 'EncryptedPowerPoint.Copy.pptx'
        $presentation = New-OfficePowerPoint -Path $sourcePath -NoSave
        try {
            Add-OfficePowerPointSlide -Presentation $presentation | Out-Null
            Save-OfficePowerPoint -Presentation $presentation -Path $savedAsPath -Password 'secret'
            Test-Path -LiteralPath $savedAsPath | Should -BeTrue
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }

        $reloaded = Get-OfficePowerPoint -Path $savedAsPath -Password 'secret'
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'does not overwrite an externally loaded encrypted presentation without a password' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.PowerPoint.PowerPointPresentation' -MethodName 'LoadEncrypted')) {
            return
        }

        $sourcePath = Join-Path $TestDrive 'EncryptedExternalPowerPoint.pptx'
        New-OfficePowerPoint -Path $sourcePath -Password 'secret' {
            PptSlide { PptTitle -Title 'Encrypted external deck' }
        }

        $presentationType = [AppDomain]::CurrentDomain.GetAssemblies() |
            ForEach-Object { $_.GetType('OfficeIMO.PowerPoint.PowerPointPresentation', $false) } |
            Where-Object { $null -ne $_ } |
            Select-Object -First 1
        $loadEncrypted = $presentationType.GetMethods() | Where-Object {
            $_.Name -eq 'LoadEncrypted' -and
            $_.GetParameters().Count -eq 3 -and
            $_.GetParameters()[0].ParameterType -eq [string] -and
            $_.GetParameters()[1].ParameterType -eq [string]
        } | Select-Object -First 1
        $presentation = $loadEncrypted.Invoke(
            $null,
            [object[]]@([string] $sourcePath, [string] 'secret', $null))
        try {
            { Save-OfficePowerPoint -Presentation $presentation -Path $sourcePath -ErrorAction Stop } |
                Should -Throw '*Provide -Password*'
            Save-OfficePowerPoint -Presentation $presentation -Path $sourcePath -Password 'secret'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation -ErrorAction SilentlyContinue
        }

        $reloaded = Get-OfficePowerPoint -Path $sourcePath -Password 'secret'
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'does not overwrite a different encrypted target from a tracked presentation without a password' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.PowerPoint.PowerPointPresentation' -MethodName 'OpenEncrypted')) {
            return
        }

        $sourcePath = Join-Path $TestDrive 'TrackedNormalPowerPoint.pptx'
        $encryptedTarget = Join-Path $TestDrive 'ProtectedPowerPointTarget.pptx'
        New-OfficePowerPoint -Path $encryptedTarget -Password 'secret' {
            PptSlide { PptTitle -Title 'Protected target' }
        }

        $presentation = New-OfficePowerPoint -Path $sourcePath -NoSave
        try {
            Add-OfficePowerPointSlide -Presentation $presentation | Out-Null
            { Save-OfficePowerPoint -Presentation $presentation -Path $encryptedTarget -ErrorAction Stop } |
                Should -Throw '*Provide -Password*'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation -ErrorAction SilentlyContinue
        }

        $reloaded = Get-OfficePowerPoint -Path $encryptedTarget -Password 'secret'
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'loads owned inspection sources as explicit read-only presentations' {
        $path = Join-Path $TestDrive 'PowerPointReadOnly.pptx'
        New-OfficePowerPoint -Path $path { PptSlide { PptTitle -Title 'Read only' } } | Out-Null

        $presentation = [PSWriteOffice.Services.PowerPoint.PowerPointDocumentService]::LoadPresentation($path, $null, $true)
        try {
            $presentation.AccessMode.ToString() | Should -Be 'ReadOnly'
            $presentation.PersistenceMode.ToString() | Should -Be 'Explicit'
        } finally {
            [PSWriteOffice.Services.PowerPoint.PowerPointDocumentService]::ClosePresentation($presentation, $false, $false)
        }
    }

    It 'preserves the existing PowerPoint service method signatures' {
        $serviceType = [PSWriteOffice.Services.PowerPoint.PowerPointDocumentService]
        $load = $serviceType.GetMethods() | Where-Object {
            $_.Name -eq 'LoadPresentation' -and
            $_.GetParameters().Count -eq 2 -and
            $_.GetParameters()[0].ParameterType -eq [string] -and
            $_.GetParameters()[1].ParameterType -eq [string]
        }
        $save = $serviceType.GetMethods() | Where-Object {
            $_.Name -eq 'SavePresentation' -and
            $_.GetParameters().Count -eq 3 -and
            $_.GetParameters()[1].ParameterType -eq [bool] -and
            $_.GetParameters()[2].ParameterType -eq [string]
        }
        $close = $serviceType.GetMethods() | Where-Object {
            $_.Name -eq 'ClosePresentation' -and
            $_.GetParameters().Count -eq 4
        }
        $closeWithPath = $serviceType.GetMethods() | Where-Object {
            $_.Name -eq 'ClosePresentation' -and
            $_.GetParameters().Count -eq 5
        }

        $load | Should -HaveCount 1
        $save | Should -HaveCount 1
        $save.ReturnType | Should -Be ([void])
        $close | Should -HaveCount 1
        $closeWithPath | Should -HaveCount 1
    }

    It 'round-trips encrypted presentations through lifecycle cmdlets' {
        if (-not (Test-OfficeLoadedMethod -TypeName 'OfficeIMO.PowerPoint.PowerPointPresentation' -MethodName 'OpenEncrypted')) {
            (Get-Command New-OfficePowerPoint).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Save-OfficePowerPoint).Parameters.Keys | Should -Contain 'Password'
            (Get-Command Get-OfficePowerPoint).Parameters.Keys | Should -Contain 'Password'
            return
        }

        $path = Join-Path $TestDrive 'EncryptedPowerPoint.pptx'

        New-OfficePowerPoint -Path $path -Password 'secret' {
            PptSlide {
                PptTitle -Title 'Encrypted deck'
            }
        }

        { Get-ZipEntriesLocal -Path $path } | Should -Throw

        $reloaded = Get-OfficePowerPoint -Path $path -Password 'secret'
        try {
            $reloaded.Slides.Count | Should -Be 1
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'creates a presentation with shapes, tables, media, and notes' {
        $path = Join-Path $TestDrive 'PowerPointContent.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 0
        $layoutType = $layouts | Where-Object { $_.LayoutType } | Select-Object -First 1
        if ($layoutType) {
            $layoutMaster = $layoutType.MasterIndex
            $layoutIndex = $layoutType.LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layoutType.LayoutType -Master $layoutType.MasterIndex -PassThru
        } elseif ($layouts[0].Name) {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layouts[0].Name -Master $layouts[0].MasterIndex -PassThru
        } else {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout $layouts[0].LayoutIndex -Master $layouts[0].MasterIndex -PassThru
        }

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Status Update'
        $shape = Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 40 -Y 40 -Width 200 -Height 80 -FillColor '#DDEEFF' -OutlineColor '#1F4E79' -OutlineWidth 1 -PassThru

        $rows = @(
            [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
            [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
        )
        $table = Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 40 -Y 140 -Width 360 -Height 200 -PassThru
        $image = Add-OfficePowerPointImage -Slide $slide -Path $imagePath -X 420 -Y 40 -Width 120 -Height 90 -PassThru
        $bullets = Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next Steps' -X 420 -Y 150 -Width 250 -Height 200 -PassThru
        Set-OfficePowerPointNotes -Slide $slide -Text 'Keep this under five minutes.'

        $slide.Shapes.Count | Should -BeGreaterThan 0
        @($slide.Tables).Count | Should -BeGreaterThan 0
        $table.Rows | Should -BeGreaterThan 0
        $shape | Should -Not -BeNullOrEmpty
        $slide.Pictures.Count | Should -BeGreaterThan 0
        $image | Should -Not -BeNullOrEmpty
        $bullets.Paragraphs.Count | Should -Be 3
        $bullets.Paragraphs[0].Text | Should -Be 'Wins'
        $notes = Get-OfficePowerPointNotes -Slide $slide
        $notes.Text | Should -Be 'Keep this under five minutes.'
        $notes.HasNotes | Should -BeTrue
        $slideSummary = Get-OfficePowerPointSlideSummary -Slide $slide
        $slideSummary.Title | Should -Be 'Status Update'
        $slideSummary.HasNotes | Should -BeTrue
        $slideSummary.NotesText | Should -Be 'Keep this under five minutes.'
        $slideSummary.ShapeCount | Should -BeGreaterThan 0
        $slideSummary.PictureCount | Should -Be 1
        $slideSummary.TableCount | Should -Be 1
        $slideSummary.PlaceholderCount | Should -BeGreaterThan 0
        $slideSummary.LayoutPlaceholderCount | Should -BeGreaterThan 0
        $shapeInfo = Get-OfficePowerPointShape -Slide $slide
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo | Where-Object Kind -eq 'Picture' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'Table' | Should -HaveCount 1
        $shapeInfo | Where-Object Kind -eq 'AutoShape' | Should -HaveCount 1
        ($shapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).MimeType | Should -Be 'image/bmp'
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).RowCount | Should -Be 3
        ($shapeInfo | Where-Object Kind -eq 'Table' | Select-Object -First 1).ColumnCount | Should -Be 2
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        $placeholder.Text | Should -Be 'Status Update'
        $placeholderUpdate = Set-OfficePowerPointPlaceholderText -Slide $slide -PlaceholderType Title -Text 'Status Update v2' -PassThru
        $placeholderUpdate.Text | Should -Be 'Status Update v2'
        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $slide2 | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $presentation

        Test-Path $path | Should -BeTrue
        $presentation.Dispose()
        $presentation = $null

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 2

            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType Title
            if (-not $reloadedPlaceholder) {
                $reloadedPlaceholder = Get-OfficePowerPointPlaceholder -Slide $reloadedSlide -PlaceholderType CenteredTitle
            }
            $reloadedPlaceholder.Text | Should -Be 'Status Update v2'
            $reloadedNotes = Get-OfficePowerPointNotes -Slide $reloadedSlide
            $reloadedNotes.Text | Should -Be 'Keep this under five minutes.'
            $reloadedNotes.SlideIndex | Should -Be 1
            $reloadedSummary = Get-OfficePowerPointSlideSummary -Presentation $reloaded -Index 1
            $reloadedSummary.Title | Should -Be 'Status Update v2'
            $reloadedSummary.SlideIndex | Should -Be 1
            $reloadedSummary.PictureCount | Should -Be 1
            $reloadedSummary.TableCount | Should -Be 1
            $reloadedShapeInfo = Get-OfficePowerPointShape -Presentation $reloaded -Index 1
            $reloadedShapeInfo.Count | Should -BeGreaterThan 0
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Table').Count | Should -Be 1
            ($reloadedShapeInfo | Where-Object Kind -eq 'Picture' | Select-Object -First 1).SlideIndex | Should -Be 1
            @($reloadedSlide.Tables).Count | Should -BeGreaterThan 0
            $reloadedSlide.Pictures.Count | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports transposed PowerPoint tables' {
        $path = Join-Path $TestDrive 'PowerPointTransposedTable.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            $rows = @(
                [PSCustomObject]@{ Region = 'Europe'; Revenue = 21704714 }
                [PSCustomObject]@{ Region = 'Asia'; Revenue = 8774099 }
            )

            Add-OfficePowerPointTable -Slide $slide -InputObject $rows -View Transpose -X 40 -Y 120 -Width 420 -Height 160 | Out-Null

            $shapeInfo = @(Get-OfficePowerPointShape -Slide $slide | Where-Object Kind -eq 'Table')[0]
            $shapeInfo.RowCount | Should -Be 3
            $shapeInfo.ColumnCount | Should -Be 3
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }
    }

    It 'preserves the legacy PowerPoint table headers alias' {
        $path = Join-Path $TestDrive 'PowerPointHeadersAlias.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            $rows = @(
                [PSCustomObject]@{ Item = 'Alpha'; Qty = 10 }
                [PSCustomObject]@{ Item = 'Beta'; Qty = 20 }
            )

            Add-OfficePowerPointTable -Slide $slide -Data $rows -Headers Qty, Item -X 40 -Y 120 -Width 420 -Height 160 | Out-Null

            $shapeInfo = @(Get-OfficePowerPointShape -Slide $slide | Where-Object Kind -eq 'Table')[0]
            $shapeInfo.RowCount | Should -Be 3
            $shapeInfo.ColumnCount | Should -Be 2
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }
    }

    It 'finds and modifies existing PowerPoint text boxes and tables' {
        $path = Join-Path $TestDrive 'PowerPointExistingShapeModify.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Add-OfficePowerPointTextBox -Slide $slide -Text 'Draft status' -X 80 -Y 80 -Width 300 -Height 50 | Out-Null
            $rows = @(
                [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
                [PSCustomObject]@{ Metric = 'Quality'; State = 'Ready' }
            )
            Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 80 -Y 160 -Width 420 -Height 180 | Out-Null

            $textShape = Find-OfficePowerPointShape -Presentation $presentation -Text 'Draft status' -Kind TextBox |
                Set-OfficePowerPointShapeText -Text 'Ready status' -PassThru
            $textShape.Text | Should -Be 'Ready status'

            $tableShape = Find-OfficePowerPointShape -Presentation $presentation -Text 'Risk' -Kind Table | Select-Object -First 1
            $tableShape | Should -Not -BeNullOrEmpty
            $row = $tableShape | Add-OfficePowerPointTableRow -Values 'Latency', 'Investigating' -PassThru
            $row.Cells[0].Text | Should -Be 'Latency'
            $row.Cells[1].Text | Should -Be 'Investigating'

            $cell = $tableShape | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Text 'Mitigating' -PassThru
            $cell.Text | Should -Be 'Mitigating'

            Save-OfficePowerPoint -Presentation $presentation
            $presentation.Dispose()
            $presentation = $null
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $readyShape = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Ready status' -Kind TextBox | Select-Object -First 1
            $readyShape | Should -Not -BeNullOrEmpty

            $updatedTable = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Mitigating' -Kind Table | Select-Object -First 1
            $updatedTable | Should -Not -BeNullOrEmpty
            $updatedTable.RowCount | Should -Be 4
            $latencyTable = Find-OfficePowerPointShape -Presentation $reloaded -Text 'Latency' -Kind Table | Select-Object -First 1
            $latencyTable | Should -Not -BeNullOrEmpty
        } finally {
            Close-OfficePowerPoint -Presentation $reloaded
        }
    }

    It 'creates and updates PowerPoint text and table cells from rich text runs' {
        foreach ($name in 'PptNew', 'PowerPointTextRun', 'PptTextRun') {
            Get-Command $name | Should -Not -BeNullOrEmpty
        }

        $path = Join-Path $TestDrive 'PowerPointRichTextRuns.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            $textBox = PptTextBox -Slide $slide -Run @(
                PptTextRun 'Status: '
                PptTextRun 'Ready' -Color SeaGreen -Bold
            ) -X 80 -Y 80 -Width 300 -Height 50 -PassThru
            $textBox.Text | Should -Be 'Status: Ready'
            $textBox = $textBox | Set-OfficePowerPointShapeText -Run @(
                PptTextRun 'Linked' -Color Crimson -BackgroundColor Yellow -FontName 'Arial' -FontSize 18 -LinkUri 'https://example.org/ppt'
            ) -PassThru
            $textBox.Paragraphs[0].Runs[0].Color | Should -Be 'DC143C'
            $textBox.Paragraphs[0].Runs[0].HighlightColor | Should -Be 'FFFF00'
            $textBox.Paragraphs[0].Runs[0].Hyperlink.AbsoluteUri | Should -Be 'https://example.org/ppt'
            $textBox = $textBox | Set-OfficePowerPointShapeText -Run @(
                PptTextRun 'Plain' -BackgroundColor None
            ) -PassThru
            $textBox.Text | Should -Be 'Plain'
            $textBox.Paragraphs[0].Runs[0].Color | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].HighlightColor | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].FontName | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].FontSize | Should -BeNullOrEmpty
            $textBox.Paragraphs[0].Runs[0].Hyperlink | Should -BeNullOrEmpty

            $table = PptTable -Slide $slide -InputObject @(
                , @(
                    @{
                        Run = @(
                            PptTextRun 'Build '
                            PptTextRun 'Ready' -Color SeaGreen
                        )
                        ColumnSpan = 2
                        FillColor = 'AliceBlue'
                        Bold = $true
                        Color = 'Red'
                        FontSize = 18
                    }
                )
                , @('Owner', 'Platform')
                , @(@{ Text = 'Filled'; FillColor = 'Yellow'; Bold = $true }, 'Styled')
                , @(@{ Run = 'Queued' }, 'String run')
            ) -X 80 -Y 150 -Width 420 -Height 140 -PassThru
            $table.GetCell(0, 0).Text | Should -Be 'Build Ready'
            $table.GetCell(0, 0).Runs[0].Bold | Should -BeTrue
            $table.GetCell(0, 0).Runs[0].Color | Should -Be 'FF0000'
            $table.GetCell(0, 0).Runs[0].FontSize | Should -Be 18
            $table.GetCell(0, 0).Runs[1].Bold | Should -BeTrue
            $table.GetCell(0, 0).Runs[1].Color | Should -Be '2E8B57'
            $table.GetCell(1, 0).Text | Should -Be 'Owner'
            $table.GetCell(2, 0).Text | Should -Be 'Filled'
            $table.GetCell(2, 0).Runs[0].Bold | Should -BeTrue
            $table.GetCell(2, 0).Runs[0].HighlightColor | Should -BeNullOrEmpty
            $table.GetCell(2, 0).FillColor | Should -Be 'FFFF00'
            $table.GetCell(3, 0).Text | Should -Be 'Queued'

            $row = $table | Add-OfficePowerPointTableRow -Values @(
                @{ Run = @(PptTextRun 'Latency '; PptTextRun 'Ready' -Color SeaGreen -Bold) },
                'SRE'
            ) -PassThru
            $row.GetCell(0).Text | Should -Be 'Latency Ready'

            $spannedRow = $table | Add-OfficePowerPointTableRow -Values @(
                @{ Text = 'Total'; ColumnSpan = 2; FillColor = 'AliceBlue' }
            ) -PassThru
            $spannedRow.GetCell(0).Text | Should -Be 'Total'
            $spannedRow.GetCell(0).Merge.Item1 | Should -Be 1
            $spannedRow.GetCell(0).Merge.Item2 | Should -Be 2

            $threeColumnTable = PptTable -Slide $slide -InputObject @(
                , @('Metric', 'Scope', 'Value')
            ) -X 80 -Y 320 -Width 420 -Height 120 -PassThru
            $valueAfterSpan = $threeColumnTable | Add-OfficePowerPointTableRow -Values @(
                @{ Text = 'Total'; ColumnSpan = 2; FillColor = 'AliceBlue' },
                '42'
            ) -PassThru
            $valueAfterSpan.GetCell(0).Merge.Item2 | Should -Be 2
            $valueAfterSpan.GetCell(2).Text | Should -Be '42'

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Values @(
                    @{ Text = 'Too wide'; ColumnSpan = 4 }
                )
            } | Should -Throw '*ColumnSpan 4*'
            $threeColumnTable.Rows | Should -Be 2

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Values @(
                    @{ Text = 'Too tall'; RowSpan = 2 }
                )
            } | Should -Throw '*RowSpan*'
            $threeColumnTable.Rows | Should -Be 2

            {
                $threeColumnTable | Add-OfficePowerPointTableRow -Values @(
                    @{ Run = @(PptTextRun 'site' -LinkUri 'https://example.org/table-cell') }
                )
            } | Should -Throw '*do not support hyperlinks yet*'
            $threeColumnTable.Rows | Should -Be 2

            $cell = $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                PptTextRun 'Owner '
                PptTextRun 'Ready' -Color Navy -Bold
            ) -PassThru
            $cell.Text | Should -Be 'Owner Ready'
            $cell = $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                PptTextRun 'Plain'
            ) -PassThru
            $cell.Text | Should -Be 'Plain'
            $cell.Runs[0].Color | Should -BeNullOrEmpty

            { $table | Set-OfficePowerPointTableCell -Row 1 -Column 1 -Run @(
                PptTextRun 'site' -LinkUri 'https://example.org/table-cell'
            ) } | Should -Throw
            $table.GetCell(1, 1).Text | Should -Be 'Plain'

            $baselineTextBox = PptTextBox -Slide $slide -Run @(
                PptTextRun 'x'
                PptTextRun '2' -Kind Superscript
            ) -X 360 -Y 80 -Width 120 -Height 50 -PassThru
            $baselineTextBox.Text | Should -Be 'x2'

            $baselineTable = PptTable -Slide $slide -InputObject @(
                , @(
                    @{
                        Run = @(
                            PptTextRun 'H'
                            PptTextRun '2' -Baseline Subscript
                            PptTextRun 'O'
                        )
                    }
                )
            ) -X 80 -Y 450 -Width 240 -Height 80 -PassThru
            $baselineTable.GetCell(0, 0).Text | Should -Be 'H2O'

            Save-OfficePowerPoint -Presentation $presentation
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }

        $slideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager($slideXml.NameTable)
        $namespaceManager.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $slideXml.SelectSingleNode('//a:r[a:t="2"]/a:rPr[@baseline="30000"]', $namespaceManager) | Should -Not -BeNullOrEmpty
        $slideXml.SelectSingleNode('//a:r[a:t="2"]/a:rPr[@baseline="-25000"]', $namespaceManager) | Should -Not -BeNullOrEmpty
    }

    It 'keeps ordinary style-named columns in PowerPoint object tables' {
        $path = Join-Path $TestDrive 'PowerPointOrdinaryStyleColumns.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            $table = PptTable -Slide $slide -InputObject @(
                [pscustomobject]@{
                    Text     = 'Apple'
                    Color    = 'Red'
                    FontSize = 'Large'
                }
            ) -X 80 -Y 150 -Width 300 -Height 100 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Text'
            $table.GetCell(0, 1).Text | Should -Be 'Color'
            $table.GetCell(0, 2).Text | Should -Be 'FontSize'
            $table.GetCell(1, 0).Text | Should -Be 'Apple'
            $table.GetCell(1, 1).Text | Should -Be 'Red'
            $table.GetCell(1, 2).Text | Should -Be 'Large'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }
    }

    It 'validates structured PowerPoint table runs before creating the table' {
        $path = Join-Path $TestDrive 'PowerPointStructuredTableRunValidation.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            @($slide.Tables).Count | Should -Be 0

            {
                PptTable -Slide $slide -InputObject @(
                    , @(
                        @{ Run = @(PptTextRun 'site' -LinkUri 'https://example.org/table-cell') }
                    )
                ) -X 80 -Y 150 -Width 300 -Height 100 -ErrorAction Stop
            } | Should -Throw '*do not support hyperlinks yet*'

            @($slide.Tables).Count | Should -Be 0
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }
    }

    It 'validates PowerPoint text box runs before creating the shape' {
        $path = Join-Path $TestDrive 'PowerPointTextBoxRunValidation.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            @($slide.TextBoxes).Count | Should -Be 0

            {
                PptTextBox -Slide $slide -Run @(
                    PptTextRun 'jump' -LinkDestinationName Summary
                ) -ErrorAction Stop
            } | Should -Throw '*named PDF/Word destinations*'

            @($slide.TextBoxes).Count | Should -Be 0
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }
    }

    It 'preserves explicit headers on structured PowerPoint tables' {
        $path = Join-Path $TestDrive 'PowerPointStructuredHeaders.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            $table = PptTable -Slide $slide -Headers Qty, Item -InputObject @(
                [pscustomobject]@{
                    Item = 'Alpha'
                    Qty  = 10
                }
                , @(
                    @{ Run = @(PptTextRun 'Total' -Bold) },
                    'Two'
                )
            ) -X 80 -Y 150 -Width 420 -Height 140 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Qty'
            $table.GetCell(0, 1).Text | Should -Be 'Item'
            $table.GetCell(1, 0).Text | Should -Be '10'
            $table.GetCell(1, 1).Text | Should -Be 'Alpha'
            $table.GetCell(2, 0).Runs[0].Text | Should -Be 'Total'
            $table.GetCell(2, 1).Text | Should -Be 'Two'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }
    }

    It 'keeps ordinary Run columns in PowerPoint object tables' {
        $path = Join-Path $TestDrive 'PowerPointOrdinaryRunColumn.pptx'
        $presentation = PptNew -Path $path -NoSave
        try {
            $slide = PptSlide -Presentation $presentation -Layout 1 -PassThru
            $table = PptTable -Slide $slide -InputObject @(
                [pscustomobject]@{
                    Run = @('Nightly', 'Daily')
                }
            ) -X 80 -Y 150 -Width 300 -Height 100 -PassThru

            $table.GetCell(0, 0).Text | Should -Be 'Run'
            $table.GetCell(1, 0).Text | Should -Match 'Nightly'
            $table.GetCell(1, 0).Text | Should -Match 'Daily'
        } finally {
            if ($presentation) {
                Close-OfficePowerPoint -Presentation $presentation
            }
        }
    }

    It 'finds existing PowerPoint shapes by metadata without a text term' {
        $path = Join-Path $TestDrive 'PowerPointMetadataShapeFind.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            $textBox = Add-OfficePowerPointTextBox -Slide $slide -Text 'Metadata status' -X 80 -Y 80 -Width 300 -Height 50 -PassThru
            $textBox.Name = 'Status.Primary'
            $rows = @(
                [PSCustomObject]@{ Metric = 'Risk'; State = 'Open' }
            )
            $table = Add-OfficePowerPointTable -Slide $slide -InputObject $rows -X 80 -Y 160 -Width 420 -Height 120 -PassThru
            $table.Name = 'Status.Table'

            $byName = Find-OfficePowerPointShape -Presentation $presentation -Name 'Status.*' -Kind TextBox
            $byName | Should -HaveCount 1
            $byName[0].Name | Should -Be 'Status.Primary'
            $byName[0].Kind | Should -Be 'TextBox'

            $byKind = Find-OfficePowerPointShape -Presentation $presentation -Kind Table
            $byKind | Should -HaveCount 1
            $byKind[0].Name | Should -Be 'Status.Table'
            $byKind[0].Kind | Should -Be 'Table'
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }
    }

    It 'arranges PowerPoint shapes through OfficeIMO layout helpers' {
        $path = Join-Path $TestDrive 'PowerPointShapeLayout.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        try {
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            $shape1 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.One' -ShapeType Rectangle -X 40 -Y 80 -Width 80 -Height 40 -PassThru
            $shape2 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.Two' -ShapeType Rectangle -X 180 -Y 140 -Width 80 -Height 40 -PassThru
            $shape3 = Add-OfficePowerPointShape -Slide $slide -Name 'Kpi.Three' -ShapeType Rectangle -X 320 -Y 200 -Width 80 -Height 40 -PassThru

            @($shape1, $shape2, $shape3) | Set-OfficePowerPointShapeLayout -Slide $slide -Align Top
            $shape1.TopPoints | Should -Be $shape2.TopPoints
            $shape1.TopPoints | Should -Be $shape3.TopPoints

            Find-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
                Set-OfficePowerPointShapeLayout -Grid -Columns 3 -Rows 1 -ToSlide -GutterXPoints 12 -NoResize | Out-Null

            $arranged = @(Get-OfficePowerPointShape -Slide $slide -Kind AutoShape | Sort-Object LeftPoints)
            $arranged | Should -HaveCount 3
            $arranged[0].LeftPoints | Should -BeLessThan $arranged[1].LeftPoints
            $arranged[1].LeftPoints | Should -BeLessThan $arranged[2].LeftPoints
        } finally {
            Close-OfficePowerPoint -Presentation $presentation
        }
    }

    It 'reads notes without creating empty notes parts' {
        $path = Join-Path $TestDrive 'PowerPointNotesRead.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'No notes yet'

        $notes = Get-OfficePowerPointNotes -Slide $slide -IncludeEmpty
        $notes.HasNotes | Should -BeFalse
        $notes.Text | Should -BeNullOrEmpty
        $notes.SlideIndex | Should -Be 0

        $presentationNotes = Get-OfficePowerPointNotes -Presentation $presentation -IncludeEmpty
        $presentationNotes.Count | Should -Be 1
        $presentationNotes[0].HasNotes | Should -BeFalse
        $slideSummary = Get-OfficePowerPointSlideSummary -Presentation $presentation
        $slideSummary.Count | Should -Be 1
        $slideSummary[0].HasNotes | Should -BeFalse
        $slideSummary[0].Title | Should -Be 'No notes yet'

        $shapeInfo = Get-OfficePowerPointShape -Presentation $presentation -Index 0 -Kind TextBox
        $shapeInfo.Count | Should -BeGreaterThan 0
        $shapeInfo[0].SlideIndex | Should -Be 0
        $shapeByIndex = Get-OfficePowerPointShape -Slide $slide -ShapeIndex $shapeInfo[0].ShapeIndex
        $shapeByIndex.ShapeIndex | Should -Be $shapeInfo[0].ShapeIndex
        $shapeByIndex.Kind | Should -Be 'TextBox'

        Save-OfficePowerPoint -Presentation $presentation

        $entries = Get-ZipEntriesLocal -Path $path
        ($entries | Where-Object { $_ -like 'ppt/notesSlides/*' }).Count | Should -Be 0
    }

    It 'persists layout placeholder edits across save and reopen' {
        $path = Join-Path $TestDrive 'PowerPointLayoutEdits.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        $layoutPlaceholder = $null

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 0

        $layoutType = $layouts | Where-Object { $_.LayoutType } | Select-Object -First 1
        if ($layoutType) {
            $layoutMaster = $layoutType.MasterIndex
            $layoutIndex = $layoutType.LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutType $layoutType.LayoutType -Master $layoutType.MasterIndex -PassThru
        } elseif ($layouts[0].Name) {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -LayoutName $layouts[0].Name -Master $layouts[0].MasterIndex -PassThru
        } else {
            $layoutMaster = $layouts[0].MasterIndex
            $layoutIndex = $layouts[0].LayoutIndex
            $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout $layouts[0].LayoutIndex -Master $layouts[0].MasterIndex -PassThru
        }

        $layoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $slide
        $layoutPlaceholders.Count | Should -BeGreaterThan 0

        $layoutPlaceholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
        if ($layoutPlaceholder) {
            $layoutPlaceholderType = $layoutPlaceholder.PlaceholderType.ToString()
            $boundsBox = Set-OfficePowerPointLayoutPlaceholderBounds -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 48 -Top 36 -Width 620 -Height 180 -PassThru
            $boundsBox.LeftPoints | Should -Be 48
            $boundsBox.TopPoints | Should -Be 36
            $boundsBox.WidthPoints | Should -Be 620
            $boundsBox.HeightPoints | Should -Be 180

            $marginsBox = Set-OfficePowerPointLayoutPlaceholderTextMargins -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Left 12 -Top 8 -Right 12 -Bottom 8 -PassThru
            $marginsBox.TextMarginLeftPoints | Should -Be 12
            $marginsBox.TextMarginTopPoints | Should -Be 8
            $marginsBox.TextMarginRightPoints | Should -Be 12
            $marginsBox.TextMarginBottomPoints | Should -Be 8

            $styleBox = Set-OfficePowerPointLayoutPlaceholderTextStyle -Presentation $presentation -Master $layoutMaster -Layout $layoutIndex -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex -Style Body -FontSize 18 -Bold $true -PassThru
            $styleBox.FontSize | Should -Be 18
            $styleBox.Bold | Should -BeTrue
        }

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedLayoutPlaceholders = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide
            if ($layoutPlaceholder) {
                $reloadedLayoutPlaceholders.Count | Should -BeGreaterThan 0
                $reloadedLayoutPlaceholder = Get-OfficePowerPointLayoutPlaceholder -Slide $reloadedSlide -PlaceholderType $layoutPlaceholderType -Index $layoutPlaceholder.PlaceholderIndex
                $reloadedLayoutPlaceholder | Should -Not -BeNullOrEmpty
                $reloadedLayoutPlaceholder.Bounds | Should -Not -BeNullOrEmpty
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.LeftPoints - 48) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.TopPoints - 36) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.WidthPoints - 620) | Should -BeLessThan 0.1
                [math]::Abs($reloadedLayoutPlaceholder.Bounds.HeightPoints - 180) | Should -BeLessThan 0.1
            }
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'removes slides and preserves the remaining title' {
        $path = Join-Path $TestDrive 'PowerPointRemoval.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave

        $firstSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $firstSlide -Title 'Remove me'

        $secondSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $secondSlide -Title 'Keep me'
        Set-OfficePowerPointPlaceholderText -Slide $secondSlide -PlaceholderType Title -Text 'Keep me v2'

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            Remove-OfficePowerPointSlide -Presentation $reloaded -Index 0 -Confirm:$false
            Save-OfficePowerPoint -Presentation $reloaded
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }

        $afterRemoval = Get-OfficePowerPoint -Path $path
        try {
            $afterRemoval.Slides.Count | Should -Be 1
            $remainingSlide = Get-OfficePowerPointSlide -Presentation $afterRemoval -Index 0
            $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType Title
            if (-not $remainingPlaceholder) {
                $remainingPlaceholder = Get-OfficePowerPointPlaceholder -Slide $remainingSlide -PlaceholderType CenteredTitle
            }
            $remainingPlaceholder.Text | Should -Be 'Keep me v2'
        } finally {
            if ($afterRemoval) {
                $afterRemoval.Dispose()
            }
        }
    }

    It 'copies a slide and preserves its content' {
        $path = Join-Path $TestDrive 'PowerPointCopy.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave

        $slide1 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Quarterly Overview'
        Add-OfficePowerPointTextBox -Slide $slide1 -Text 'Revenue and margin summary' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $slide1 -Text 'Reuse this for the board deck.'

        $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Closing Slide'

        $copiedSlide = Copy-OfficePowerPointSlide -Presentation $presentation -Index 0 -InsertAt 1 -PassThru
        $copiedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 3

            $copied = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType Title
            if (-not $copiedTitle) {
                $copiedTitle = Get-OfficePowerPointPlaceholder -Slide $copied -PlaceholderType CenteredTitle
            }

            $copiedTitle.Text | Should -Be 'Quarterly Overview'
            (Get-OfficePowerPointNotes -Slide $copied).Text | Should -Be 'Reuse this for the board deck.'
            ($copied.Shapes.Count -gt 0) | Should -BeTrue

            $lastSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 2
            $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType Title
            if (-not $lastTitle) {
                $lastTitle = Get-OfficePowerPointPlaceholder -Slide $lastSlide -PlaceholderType CenteredTitle
            }

            $lastTitle.Text | Should -Be 'Closing Slide'
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'sets slide transitions and custom slide sizes' {
        $path = Join-Path $TestDrive 'PowerPointTransitionsAndSize.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Transition Demo' | Out-Null

        $updatedSlide = $slide | Set-OfficePowerPointSlideTransition -Transition Fade -PassThru
        $fadeTransition = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.PowerPoint' -TypeName 'OfficeIMO.PowerPoint.PowerPointSlideTransition' -Name 'Fade' -CommandName 'New-OfficePowerPoint'
        $updatedSlide.Transition | Should -Be $fadeTransition

        $slideSize = Set-OfficePowerPointSlideSize -Presentation $presentation -WidthCm 25.4 -HeightCm 14.0 -PassThru
        [math]::Round($slideSize.WidthCm, 1) | Should -Be 25.4
        [math]::Round($slideSize.HeightCm, 1) | Should -Be 14.0

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $slideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $transitionNode = $slideXml.SelectSingleNode("/*[local-name()='sld']/*[local-name()='transition']")
        $transitionNode | Should -Not -BeNullOrEmpty
        $transitionNode.SelectSingleNode("*[local-name()='fade']") | Should -Not -BeNullOrEmpty

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedSlide.Transition | Should -Be $fadeTransition
            [math]::Round($reloaded.SlideSize.WidthCm, 1) | Should -Be 25.4
            [math]::Round($reloaded.SlideSize.HeightCm, 1) | Should -Be 14.0
            $reloaded.SlideSize.IsLandscape | Should -BeTrue
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'applies preset slide sizes including portrait orientation' {
        $path = Join-Path $TestDrive 'PowerPointPresetSize.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave
        Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 | Out-Null

        $presetSize = Set-OfficePowerPointSlideSize -Presentation $presentation -Preset Screen4x3 -Portrait -PassThru
        $presetSize.IsPortrait | Should -BeTrue
        [math]::Round($presetSize.WidthInches, 1) | Should -Be 7.5
        [math]::Round($presetSize.HeightInches, 1) | Should -Be 10.0

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.SlideSize.IsPortrait | Should -BeTrue
            [math]::Round($reloaded.SlideSize.WidthInches, 1) | Should -Be 7.5
            [math]::Round($reloaded.SlideSize.HeightInches, 1) | Should -Be 10.0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'creates a presentation via the DSL context' {
        $path = Join-Path $TestDrive 'PowerPointDsl.pptx'

        New-OfficePowerPoint -Path $path {
            $layout = @(Get-OfficePowerPointLayout)[0]
            if ($layout) {
                Set-OfficePowerPointLayoutPlaceholderBounds -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Left 60 -Top 40 -Width 600 -Height 120 -CreateIfMissing
                Set-OfficePowerPointLayoutPlaceholderTextMargins -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Left 8 -Top 6 -Right 8 -Bottom 6 -CreateIfMissing
                Set-OfficePowerPointLayoutPlaceholderTextStyle -Master $layout.MasterIndex -Layout $layout.LayoutIndex -PlaceholderType Title `
                    -Style Title -FontSize 32 -Bold $true -CreateIfMissing
            }

            PptSlide {
                PptTitle -Title 'DSL Slide'
                PptTextBox -Text 'Hello from DSL' -X 80 -Y 150 -Width 240 -Height 60
                $layoutPlaceholders = PptLayoutPlaceholders
                $layoutPlaceholders.Count | Should -BeGreaterThan 0
                $layoutPlaceholder = $layoutPlaceholders | Where-Object { $_.PlaceholderType } | Select-Object -First 1
                if ($layoutPlaceholder) {
                    $placeholderType = $layoutPlaceholder.PlaceholderType.ToString()
                    $placeholder = Get-OfficePowerPointPlaceholder -PlaceholderType $placeholderType -Index $layoutPlaceholder.PlaceholderIndex
                    if ($placeholder -and ((
                            $placeholderType -eq 'Title'
                        ) -or (
                            $placeholderType -eq 'CenteredTitle'
                        ))) {
                        $placeholder | Should -Not -BeNullOrEmpty
                        $placeholder.Text | Should -Be 'DSL Slide'
                    }
                }
            }
        }

        Test-Path $path | Should -BeTrue

        $presentation = Get-OfficePowerPoint -Path $path
        $slide = Get-OfficePowerPointSlide -Presentation $presentation -Index 0
        $summary = Get-OfficePowerPointSlideSummary -Slide $slide
        $summary.LayoutPlaceholderCount | Should -BeGreaterThan 0
        if ($summary.Title) {
            $summary.Title | Should -Be 'DSL Slide'
        }
        $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
        if (-not $placeholder) {
            $placeholder = Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType CenteredTitle
        }
        if ($placeholder) {
            $placeholder.Text | Should -Be 'DSL Slide'
        }

        Save-OfficePowerPoint -Presentation $presentation
    }

    It 'supports sections, text replacement, and slide import' {
        $sourcePath = Join-Path $TestDrive 'PowerPointSourceDeck.pptx'
        $targetPath = Join-Path $TestDrive 'PowerPointTargetDeck.pptx'

        $source = New-OfficePowerPoint -Path $sourcePath -NoSave
        $sourceSlide = Add-OfficePowerPointSlide -Presentation $source -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $sourceSlide -Title 'FY24 Imported'
        Add-OfficePowerPointTextBox -Slide $sourceSlide -Text 'FY24 details from source' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $sourceSlide -Text 'FY24 source notes'
        Save-OfficePowerPoint -Presentation $source

        $target = New-OfficePowerPoint -Path $targetPath -NoSave
        $introSlide = Add-OfficePowerPointSlide -Presentation $target -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $introSlide -Title 'FY24 Overview'
        Add-OfficePowerPointTextBox -Slide $introSlide -Text 'FY24 summary for leadership' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null
        Set-OfficePowerPointNotes -Slide $introSlide -Text 'FY24 note for intro'

        $resultsSlide = Add-OfficePowerPointSlide -Presentation $target -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $resultsSlide -Title 'FY24 Results'
        Add-OfficePowerPointTextBox -Slide $resultsSlide -Text 'FY24 results body' -X 80 -Y 150 -Width 320 -Height 60 | Out-Null

        $introSection = Add-OfficePowerPointSection -Presentation $target -Name 'Intro' -StartSlideIndex 0 -PassThru
        $introSection.Name | Should -Be 'Intro'
        $resultsSection = Add-OfficePowerPointSection -Presentation $target -Name 'Results' -StartSlideIndex 1 -PassThru
        $resultsSection.Name | Should -Be 'Results'

        $renamedSection = Rename-OfficePowerPointSection -Presentation $target -Name 'Results' -NewName 'Deep Dive' -PassThru
        $renamedSection.Name | Should -Be 'Deep Dive'

        $sections = @(Get-OfficePowerPointSection -Presentation $target)
        $sections.Count | Should -Be 2
        ($sections | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
        ($sections | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 1

        $replacements = Update-OfficePowerPointText -Presentation $target -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes -PassThru
        $replacements | Should -BeGreaterThan 0

        $importedSlide = Import-OfficePowerPointSlide -Presentation $target -SourcePath $sourcePath -SourceIndex 0 -InsertAt 1
        $importedSlide | Should -Not -BeNullOrEmpty

        Save-OfficePowerPoint -Presentation $target

        $reloaded = Get-OfficePowerPoint -Path $targetPath
        try {
            $reloaded.Slides.Count | Should -Be 3

            $slide0 = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType Title
            if (-not $title0) {
                $title0 = Get-OfficePowerPointPlaceholder -Slide $slide0 -PlaceholderType CenteredTitle
            }
            $title0.Text | Should -Be 'FY25 Overview'
            (Get-OfficePowerPointNotes -Slide $slide0).Text | Should -Be 'FY25 note for intro'

            $slide1 = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType Title
            if (-not $title1) {
                $title1 = Get-OfficePowerPointPlaceholder -Slide $slide1 -PlaceholderType CenteredTitle
            }
            $title1.Text | Should -Be 'FY24 Imported'
            (Get-OfficePowerPointNotes -Slide $slide1).Text | Should -Be 'FY24 source notes'

            $sectionInfo = @(Get-OfficePowerPointSection -Presentation $reloaded)
            $sectionInfo.Count | Should -Be 2
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 0
            ($sectionInfo | Where-Object Name -eq 'Intro').SlideIndices | Should -Contain 1
            ($sectionInfo | Where-Object Name -eq 'Deep Dive').SlideIndices | Should -Contain 2

            $aliasReplacements = Replace-OfficePowerPointText -Presentation $reloaded -OldValue 'FY24' -NewValue 'FY26' -PassThru
            $aliasReplacements | Should -BeGreaterThan 0
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports theme inspection, theme updates, and slide layout switching' {
        $path = Join-Path $TestDrive 'PowerPointThemeAndLayout.pptx'
        $presentation = New-OfficePowerPoint -Path $path -NoSave

        $slide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
        Set-OfficePowerPointSlideTitle -Slide $slide -Title 'Theme Demo' | Out-Null

        $layouts = Get-OfficePowerPointLayout -Presentation $presentation
        $layouts.Count | Should -BeGreaterThan 1
        $currentLayoutIndex = $slide.LayoutIndex
        $alternativeLayout = $layouts | Where-Object { $_.LayoutIndex -ne $currentLayoutIndex } | Select-Object -First 1
        $alternativeLayout | Should -Not -BeNullOrEmpty

        $themeBefore = Get-OfficePowerPointTheme -Presentation $presentation
        $themeBefore.MasterIndex | Should -Be 0
        $themeBefore.Colors.Count | Should -BeGreaterThan 0

        Set-OfficePowerPointThemeColor -Presentation $presentation -Colors @{
            Accent1 = '#C00000'
            Accent2 = '00B0F0'
        } -AllMasters
        Set-OfficePowerPointThemeFonts -Presentation $presentation -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
        Set-OfficePowerPointThemeName -Presentation $presentation -Name 'Contoso Theme' -AllMasters

        if ($alternativeLayout.LayoutType) {
            $slide | Set-OfficePowerPointSlideLayout -LayoutType $alternativeLayout.LayoutType -Master $alternativeLayout.MasterIndex | Out-Null
        } elseif ($alternativeLayout.Name) {
            $slide | Set-OfficePowerPointSlideLayout -LayoutName $alternativeLayout.Name -Master $alternativeLayout.MasterIndex | Out-Null
        } else {
            $slide | Set-OfficePowerPointSlideLayout -Layout $alternativeLayout.LayoutIndex -Master $alternativeLayout.MasterIndex | Out-Null
        }

        $themeAfter = Get-OfficePowerPointTheme -Presentation $presentation
        $themeAfter.ThemeName | Should -Be 'Contoso Theme'
        $accent1Color = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.PowerPoint' -TypeName 'OfficeIMO.PowerPoint.PowerPointThemeColor' -Name 'Accent1' -CommandName 'New-OfficePowerPoint'
        $accent2Color = Get-TestPSWriteOfficeEnumValue -AssemblyName 'OfficeIMO.PowerPoint' -TypeName 'OfficeIMO.PowerPoint.PowerPointThemeColor' -Name 'Accent2' -CommandName 'New-OfficePowerPoint'
        $themeAfter.Colors[$accent1Color] | Should -Be 'C00000'
        $themeAfter.Colors[$accent2Color] | Should -Be '00B0F0'
        $themeAfter.MajorLatin | Should -Be 'Aptos'
        $themeAfter.MinorLatin | Should -Be 'Calibri'
        $slide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex

        Save-OfficePowerPoint -Presentation $presentation
        $presentation.Dispose()

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloadedTheme = Get-OfficePowerPointTheme -Presentation $reloaded
            $reloadedTheme.ThemeName | Should -Be 'Contoso Theme'
            $reloadedTheme.Colors[$accent1Color] | Should -Be 'C00000'
            $reloadedTheme.Colors[$accent2Color] | Should -Be '00B0F0'
            $reloadedTheme.MajorLatin | Should -Be 'Aptos'
            $reloadedTheme.MinorLatin | Should -Be 'Calibri'

            $reloadedSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedSlide.LayoutIndex | Should -Be $alternativeLayout.LayoutIndex
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'supports slide backgrounds and layout box helpers' {
        $path = Join-Path $TestDrive 'PowerPointBackgroundsAndLayout.pptx'
        $imagePath = New-TestOfficeImageFile -Directory $TestDrive
        $presentation = $null

        try {
            $presentation = New-OfficePowerPoint -Path $path -NoSave

            Set-OfficePowerPointSlideSize -Presentation $presentation -WidthCm 30 -HeightCm 20 | Out-Null

            $colorSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Set-OfficePowerPointSlideTitle -Slide $colorSlide -Title 'Layout Demo' | Out-Null
            $updatedSlide = Set-OfficePowerPointBackground -Slide $colorSlide -Color '#F4F7FB' -PassThru
            $updatedSlide | Should -Be $colorSlide
            $colorSlide.BackgroundColor | Should -Be 'f4f7fb'

            $contentBox = Get-OfficePowerPointLayoutBox -Presentation $presentation -MarginCm 1.5
            [math]::Round($contentBox.LeftCm, 2) | Should -Be 1.5
            [math]::Round($contentBox.TopCm, 2) | Should -Be 1.5
            [math]::Round($contentBox.WidthCm, 2) | Should -Be 27
            [math]::Round($contentBox.HeightCm, 2) | Should -Be 17

            $columns = @(Get-OfficePowerPointLayoutBox -Presentation $presentation -ColumnCount 2 -MarginCm 1.5 -GutterCm 1.0)
            $columns.Count | Should -Be 2
            [math]::Round($columns[0].WidthCm, 2) | Should -Be 13
            [math]::Round($columns[1].LeftCm, 2) | Should -Be 15.5

            $rows = @(Get-OfficePowerPointLayoutBox -Presentation $presentation -RowCount 2 -MarginCm 1.5 -GutterCm 0.5)
            $rows.Count | Should -Be 2
            [math]::Round($rows[0].HeightCm, 2) | Should -Be 8.25
            [math]::Round($rows[1].TopCm, 2) | Should -Be 10.25

            Add-OfficePowerPointTextBox -Slide $colorSlide -Text 'Left column' -X $columns[0].LeftPoints -Y $columns[0].TopPoints -Width $columns[0].WidthPoints -Height 40 | Out-Null
            Add-OfficePowerPointTextBox -Slide $colorSlide -Text 'Right column' -X $columns[1].LeftPoints -Y $columns[1].TopPoints -Width $columns[1].WidthPoints -Height 40 | Out-Null

            $imageSlide = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Set-OfficePowerPointSlideTitle -Slide $imageSlide -Title 'Image Background' | Out-Null
            Set-OfficePowerPointBackground -Slide $imageSlide -ImagePath $imagePath | Out-Null
            $imageSlide.BackgroundColor | Should -BeNullOrEmpty

            Save-OfficePowerPoint -Presentation $presentation
        } finally {
            if ($presentation) {
                $presentation.Dispose()
            }
        }

        $colorSlideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 0)
        $colorSlideXml.OuterXml | Should -Match 'rgbClr val="f4f7fb"'

        $imageSlideXml = Get-ZipXmlDocumentLocal -Path $path -Entry (Get-TestPowerPointSlideEntry -Path $path -Index 1)
        $imageSlideXml.OuterXml | Should -Match '<a:blip '
        $imageSlideXml.OuterXml | Should -Match 'r:embed='

        $entries = @(Get-ZipEntriesLocal -Path $path)
        ($entries | Where-Object { $_ -like 'ppt/media/*' }).Count | Should -BeGreaterThan 0

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloadedColorSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 0
            $reloadedColorSlide.BackgroundColor | Should -Be 'f4f7fb'

            $reloadedImageSlide = Get-OfficePowerPointSlide -Presentation $reloaded -Index 1
            $reloadedImageSlide.BackgroundColor | Should -BeNullOrEmpty
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'renders OfficeIMO designer deck plans through the PowerPoint DSL' {
        $path = Join-Path $TestDrive 'PowerPointDesignerDeck.pptx'
        $plan = PptDeckPlan {
            PptPlanSection -Title 'Reliability Showcase' -Subtitle 'From script to polished deck'
            PptPlanProcess -Title 'Delivery Flow' -Steps @(
                [PSCustomObject]@{ Number = '01'; Title = 'Author'; Body = 'Describe the story once.' }
                [PSCustomObject]@{ Number = '02'; Title = 'Render'; Body = 'Let OfficeIMO compose the layout.' }
                [PSCustomObject]@{ Number = '03'; Title = 'Validate'; Body = 'Inspect and test the package.' }
            )
            PptPlanCardGrid -Title 'Proof Points' -Cards @(
                [PSCustomObject]@{ Title = 'Useful'; Items = @('Semantic input', 'Readable output') }
                [PSCustomObject]@{ Title = 'Fast'; Items = @('One cmdlet bridge', 'Deterministic seed') }
            )
        }

        $preview = @(PptDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'designer-test' -Purpose 'technical service brief' -Preview -PassThru)
        $preview.Count | Should -BeGreaterThan 0

        New-OfficePowerPoint -Path $path {
            PptDesignerDeck -Plan $plan -AccentColor '#008C95' -Seed 'designer-test' -Purpose 'technical service brief' -Name 'Designer Test' -LayoutStrategy ContentFirst
        }

        $entries = @(Get-ZipEntriesLocal -Path $path)
        ($entries | Where-Object { $_ -match '^ppt/slides/slide\d+\.xml$' }).Count | Should -BeGreaterThan 2

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -BeGreaterThan 2
            $summary = @(Get-OfficePowerPointSlideSummary -Presentation $reloaded)
            $summary.Count | Should -BeGreaterThan 2
            ($summary | Where-Object { $_.ShapeCount -gt 0 }).Count | Should -BeGreaterThan 2
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }

    It 'rejects designer plan data arrays that contain only null values' {
        $mapper = [PSWriteOffice.Cmdlets.PowerPoint.AddOfficePowerPointPlanProcessCommand].Assembly.GetType('PSWriteOffice.Services.PowerPoint.PowerPointDesignerDataMapper')
        $method = $mapper.GetMethod('ToProcessSteps', [System.Reflection.BindingFlags] 'Public,Static')

        try {
            $method.Invoke($null, (, [object[]] @($null)))
            throw 'Expected mapper to reject all-null process data.'
        } catch {
            $_.Exception.InnerException.Message | Should -Be 'Process steps require at least one item.'
        }
    }

    It 'normalizes blank optional designer colors to null' {
        $mapper = [PSWriteOffice.Cmdlets.PowerPoint.AddOfficePowerPointPlanProcessCommand].Assembly.GetType('PSWriteOffice.Services.PowerPoint.PowerPointDesignerDataMapper')
        $method = $mapper.GetMethod('ToCards', [System.Reflection.BindingFlags] 'Public,Static')

        $cards = $method.Invoke($null, (, [object[]] @([PSCustomObject]@{ Title = 'Blank color'; Items = @('Still valid'); AccentColor = '   ' })))
        $cards[0].AccentColor | Should -BeNullOrEmpty
    }

    It 'supports PowerPoint charts from object data' {
        $path = Join-Path $TestDrive 'PowerPointCharts.pptx'
        $presentation = $null
        $rows = @(
            [PSCustomObject]@{ Month = 'Jan'; MonthNumber = 1; Sales = 10; Profit = 4 }
            [PSCustomObject]@{ Month = 'Feb'; MonthNumber = 2; Sales = 14; Profit = 6 }
            [PSCustomObject]@{ Month = 'Mar'; MonthNumber = 3; Sales = 18; Profit = 8 }
        )

        try {
            $presentation = New-OfficePowerPoint -Path $path -NoSave

            $slide1 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Set-OfficePowerPointSlideTitle -Slide $slide1 -Title 'Column Chart' | Out-Null
            $columnChart = Add-OfficePowerPointChart -Slide $slide1 -Data $rows -CategoryProperty Month -SeriesProperty Sales, Profit -Title 'Sales vs Profit' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
            $columnChart | Should -Not -BeNullOrEmpty
            @($slide1.Charts).Count | Should -Be 1

            $slide2 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Set-OfficePowerPointSlideTitle -Slide $slide2 -Title 'Pie Chart' | Out-Null
            $pieChart = Add-OfficePowerPointChart -Slide $slide2 -Type Pie -InputObject $rows -CategoryProperty Month -SeriesProperty Sales -Title 'Sales Mix' -X 40 -Y 120 -Width 320 -Height 220 -PassThru
            $pieChart | Should -Not -BeNullOrEmpty
            @($slide2.Charts).Count | Should -Be 1

            $slide3 = Add-OfficePowerPointSlide -Presentation $presentation -Layout 1 -PassThru
            Set-OfficePowerPointSlideTitle -Slide $slide3 -Title 'Scatter Chart' | Out-Null
            $scatterChart = Add-OfficePowerPointChart -Slide $slide3 -Type Scatter -Data $rows -XProperty MonthNumber -YProperty Sales, Profit -Title 'Trend Scatter' -X 40 -Y 120 -Width 360 -Height 220 -PassThru
            $scatterChart | Should -Not -BeNullOrEmpty
            @($slide3.Charts).Count | Should -Be 1

            (Get-OfficePowerPointShape -Slide $slide1 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1
            (Get-OfficePowerPointShape -Slide $slide2 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1
            (Get-OfficePowerPointShape -Slide $slide3 | Where-Object Kind -eq 'Chart') | Should -HaveCount 1

            Save-OfficePowerPoint -Presentation $presentation
        } finally {
            if ($presentation) {
                $presentation.Dispose()
            }
        }

        $chartEntries = @(Get-ZipEntriesLocal -Path $path | Where-Object { $_ -match '^ppt/charts/chart\d+\.xml$' } | Sort-Object)
        $chartEntries.Count | Should -Be 3

        $chart1Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[0]
        $chart1Xml.OuterXml | Should -Match '<c:barChart'
        $chart1Xml.OuterXml | Should -Match 'Sales vs Profit'

        $chart2Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[1]
        $chart2Xml.OuterXml | Should -Match '<c:pieChart'
        $chart2Xml.OuterXml | Should -Match 'Sales Mix'

        $chart3Xml = Get-ZipXmlDocumentLocal -Path $path -Entry $chartEntries[2]
        $chart3Xml.OuterXml | Should -Match '<c:scatterChart'
        $chart3Xml.OuterXml | Should -Match 'Trend Scatter'

        $reloaded = Get-OfficePowerPoint -Path $path
        try {
            $reloaded.Slides.Count | Should -Be 3
            (Get-OfficePowerPointShape -Presentation $reloaded | Where-Object Kind -eq 'Chart') | Should -HaveCount 3
        } finally {
            if ($reloaded) {
                $reloaded.Dispose()
            }
        }
    }
}
