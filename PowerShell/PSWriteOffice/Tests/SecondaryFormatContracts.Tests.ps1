BeforeAll {
    $modulePath = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '../PSWriteOffice.psd1'
    }
    Import-Module $modulePath -Global -ErrorAction Stop
}

Describe 'Secondary format native binding and authoring' {
    It 'converts physical container lengths on metric pages and preserves native text styling' {
        $document = New-OfficeVisio -DisplayUnit Centimeters
        try {
            $page = $document.Pages[0]
            $member = Add-OfficeVisioRectangle -Page $page -X 2in -Y 2in -Width 1in -Height 0.5in -Text Member -PassThru
            $container = Add-OfficeVisioContainer -Page $page -InputObject $member -Key group -Margin 0.25in -HeadingHeight 0pt -LineWeight 1pt -FillColor rebeccapurple -LineColor '#123456' -Text Group -Font Aptos -Size 12pt -Bold:$false -Italic -Underline Double -Strikethrough -Color Navy -Alignment Center -PassThru
            $container.Width.ToInches() | Should -Be 1.5
            $container.Height.ToInches() | Should -Be 1
            $container.LineWeight | Should -Be (1 / 72)
            $container.FillColor.ToString() | Should -Be '#663399'
            $container.TextStyle.FontFamily | Should -Be Aptos
            $container.TextStyle.Size | Should -Be 12
            $container.TextStyle.Bold | Should -BeFalse
            $container.TextStyle.Italic | Should -BeTrue
            $container.TextStyle.UnderlineStyle.ToString() | Should -Be Double
            $container.TextStyle.Strikethrough | Should -BeTrue
            $container.TextStyle.Color.ToString() | Should -Be '#000080'
            $container.TextStyle.HorizontalAlignment.ToString() | Should -Be Center
            $path = Join-Path $TestDrive 'metric-container.vsdx'
            Save-OfficeVisio -Document $document -Path $path
            (Get-Item -LiteralPath $path).Length | Should -BeGreaterThan 0
        } finally { Close-OfficeVisio $document }
    }

    It 'imports catalog objects and literal file properties through their native pipeline routes' {
        $catalog = Import-OfficeVisioStencilCatalog -BuiltIn Flowchart
        $piped = $catalog | Import-OfficeVisioStencilCatalog
        [object]::ReferenceEquals($catalog, $piped) | Should -BeTrue
        $document = New-OfficeVisio -UseMastersByDefault
        try {
            Add-OfficeVisioStencilShape -Page $document.Pages[0] -BuiltIn Flowchart -Stencil process -Text Pipeline
            $path = Join-Path $TestDrive 'catalog[1].vsdx'
            Save-OfficeVisio -Document $document -Path $path
            $loaded = Get-Item -LiteralPath $path | Import-OfficeVisioStencilCatalog
            $loaded.Shapes.Count | Should -BeGreaterThan 0
        } finally { Close-OfficeVisio $document }
    }

    It 'binds mutually exclusive layout modes and converts stack spacing into inches' {
        $document = New-OfficeVisio
        try {
            $page = $document.Pages[0]
            $one = Add-OfficeVisioRectangle -Page $page -Text One -X 1in -Y 2in -Width 1in -Height 1in -PassThru
            $two = Add-OfficeVisioRectangle -Page $page -Text Two -X 4in -Y 2in -Width 1in -Height 1in -PassThru
            { Set-OfficeVisioShapeLayout -InputObject $one, $two -Grid -HorizontalStack -ErrorAction Stop } |
                Should -Throw -ErrorId 'AmbiguousParameterSet,*'
            @($one, $two) | Set-OfficeVisioShapeLayout -Page $page -HorizontalStack -HorizontalSpacing 18pt -NoRouteInternalConnectors
            ($two.PinX.ToInches() - $one.PinX.ToInches()) | Should -Be 1.25
        } finally { Close-OfficeVisio $document }
    }

    It 'carries physical positions and native typography into an ODP text box' {
        $document = New-OfficeOpenDocument -Kind Presentation
        try {
            $slide = Add-OfficeOpenDocumentSlide -Document $document -Name Review -PassThru
            $box = Add-OfficeOpenDocumentTextBox -Slide $slide -Text Ready -X 72 -Y 25.4mm -Width 2in -Height 36pt -Font Aptos -Size 12pt -Bold -Italic:$false -Underline Wave -Strikethrough -Color Navy -Alignment Center -PassThru
            $box.Bounds.X.ToCentimeters() | Should -Be 2.54
            $box.Bounds.Y.ToCentimeters() | Should -Be 2.54
            $box.Bounds.Width.ToCentimeters() | Should -Be 5.08
            $box.Bounds.Height.ToCentimeters() | Should -Be 1.27
            $paragraph = $box.Paragraphs[0]
            $paragraph.FontFamily | Should -Be Aptos
            $paragraph.FontSize.ToPoints() | Should -Be 12
            $paragraph.Bold | Should -BeTrue
            $paragraph.Italic | Should -BeFalse
            $paragraph.UnderlineStyle.ToString() | Should -Be Wave
            $paragraph.Strikethrough | Should -BeTrue
            $paragraph.Color.ToString() | Should -Be '#000080'
            $paragraph.Alignment.ToString() | Should -Be Center
            $path = Join-Path $TestDrive 'positioned.odp'
            Save-OfficeOpenDocument -Document $document -Path $path
            (Get-Item -LiteralPath $path).Length | Should -BeGreaterThan 0
        } finally { $document.Dispose() }
    }

    It 'authors ODT headings and paragraphs with explicit false switches and native typography' {
        $document = New-OfficeOpenDocument -Kind Text
        try {
            $heading = Add-OfficeOpenDocumentHeading -Document $document -Text Results -Level 2 -Font Aptos -Size 14pt -Bold -PassThru
            $paragraph = Add-OfficeOpenDocumentParagraph -Document $document -Text Approved -Font Calibri -Size 12pt -Bold:$false -Italic -Underline Dash -Strikethrough -Color '#123456' -Alignment Justify -PassThru
            $heading.FontFamily | Should -Be Aptos
            $heading.FontSize.ToPoints() | Should -Be 14
            $heading.Bold | Should -BeTrue
            $paragraph.FontFamily | Should -Be Calibri
            $paragraph.FontSize.ToPoints() | Should -Be 12
            $paragraph.Bold | Should -BeFalse
            $paragraph.Italic | Should -BeTrue
            $paragraph.UnderlineStyle.ToString() | Should -Be Dash
            $paragraph.Strikethrough | Should -BeTrue
            $paragraph.Color.ToString() | Should -Be '#123456'
            $paragraph.Alignment.ToString() | Should -Be Justify
            $path = Join-Path $TestDrive 'styled.odt'
            Save-OfficeOpenDocument -Document $document -Path $path
            (Get-Item -LiteralPath $path).Length | Should -BeGreaterThan 0
        } finally { $document.Dispose() }
    }

    It 'writes native RTF fonts decorations colors and links and converts fallback image sizes to points' {
        $document = New-OfficeRtf
        try {
            Add-OfficeRtfParagraph -Document $document -Text Approved -Font Aptos -Size 12pt -Bold:$false -Italic -Underline Double -Strikethrough -Color Navy -Link 'https://example.invalid/review' -Alignment Center
            $run = $document.Paragraphs[0].Runs[0]
            $document.Fonts.Name | Should -Contain Aptos
            $run.FontSize | Should -Be 12
            $run.Bold | Should -BeFalse
            $run.Italic | Should -BeTrue
            $run.Strikethrough | Should -BeTrue
            $run.UnderlineStyle.ToString() | Should -Be Double
            $run.ForegroundColorIndex | Should -BeGreaterThan 0
            $run.Hyperlink.AbsoluteUri | Should -Be 'https://example.invalid/review'
            $path = Join-Path $TestDrive 'styled.rtf'
            Save-OfficeRtf -Document $document -Path $path
            $rtf = [IO.File]::ReadAllText($path)
            $rtf | Should -Match '\\fs24'
            $rtf | Should -Match '\\uldb'
            $rtf | Should -Match '\\strike'
            $rtf | Should -Match 'HYPERLINK'
            $options = New-OfficeRtfPdfOptions -DefaultImageWidth 1in -DefaultImageHeight 25.4mm
            $options.DefaultImageWidth | Should -Be 72
            $options.DefaultImageHeight | Should -Be 72
            { New-OfficeRtfPdfOptions -DefaultImageWidth 0pt -ErrorAction Stop } | Should -Throw '*positive*'
        } finally { Close-OfficeRtf -Document $document }
    }

    It 'plans Confluence writes offline and gates live mutations through ShouldProcess' {
        $create = Publish-OfficeConfluencePage -SpaceId 42 -Title Review -Content Ready -PlanOnly
        $create.Method | Should -Be POST
        $update = Publish-OfficeConfluencePage -PageId 12 -Title Review -VersionNumber 2 -Content Ready -PlanOnly
        $update.Method | Should -Be PUT
        ($update.Payload | ConvertFrom-Json).version.number | Should -Be 2
        (Remove-OfficeConfluencePage -PageId 12 -PlanOnly).Method | Should -Be DELETE
        $session = New-OfficeConfluenceSession -SiteUri 'https://example.invalid' -UserName synthetic@example.invalid -Password 'synthetic-token'
        { Publish-OfficeConfluencePage -Session $session -SpaceId 42 -Title Review -Content Ready -PlanOnly -ErrorAction Stop } |
            Should -Throw -ErrorId 'AmbiguousParameterSet,*'
        @(Publish-OfficeConfluencePage -Session $session -SpaceId 42 -Title Review -Content Ready -WhatIf).Count | Should -Be 0
        @(Remove-OfficeConfluencePage -Session $session -PageId 12 -WhatIf).Count | Should -Be 0
    }

    It 'binds Project task positions and leaves schedule application explicit under WhatIf' {
        $document = New-OfficeProject -Name Delivery -StartDate '2026-11-02 08:00'
        try {
            $one = Add-OfficeProjectTask -Document $document -Name One -DurationDays 1 -PassThru
            $two = Add-OfficeProjectTask -Document $document -Name Two -DurationDays 1 -PassThru
            $dependency = Add-OfficeProjectDependency $one $two -Document $document -PassThru
            $dependency | Should -Not -BeNullOrEmpty
            $resource = Add-OfficeProjectResource -Document $document -Name Worker -PassThru
            $assignment = Add-OfficeProjectAssignment $one $resource -Document $document -UnitsPercent 50 -PassThru
            $assignment.Task.Name | Should -Be One
            $assignment.Resource.Name | Should -Be Worker
            $revision = $document.Revision
            @(Invoke-OfficeProjectSchedule -Document $document -Apply -WhatIf).Count | Should -Be 0
            $document.Revision | Should -Be $revision
            $one.Start | Should -BeNullOrEmpty
            $schedule = Invoke-OfficeProjectSchedule -Document $document
            $schedule | Should -Not -BeNullOrEmpty
            Get-OfficeProjectView -Document $document -Schedule $schedule -Options @{ Kind = 'Gantt' } | Should -Not -BeNullOrEmpty
            $options = New-OfficeProjectViewOptions -Kind ResourceUsage -Timescale Month
            $options.Kind.ToString() | Should -Be ResourceUsage
            (Get-OfficeProjectView -Document $document -Options $options).Kind.ToString() |
                Should -Be ResourceUsage
            $one.Start | Should -BeNullOrEmpty
            { Get-OfficeProjectView -Document $document -Kind Gantt -Options @{ Kind = 'Gantt' } -ErrorAction Stop } |
                Should -Throw -ErrorId 'AmbiguousParameterSet,*'
        } finally { $document.Dispose() }
    }
}
