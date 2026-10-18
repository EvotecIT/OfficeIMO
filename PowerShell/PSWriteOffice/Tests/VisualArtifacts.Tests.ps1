BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop
}

Describe 'ChartForgeX visual artifacts' {
    BeforeAll {
        $script:SvgPath = Join-Path $TestDrive 'service-health.svg'
        $svg = @'
<svg xmlns="http://www.w3.org/2000/svg" width="400" height="220" viewBox="0 0 400 220">
  <rect width="400" height="220" rx="16" fill="#f8fafc"/>
  <text x="24" y="42" font-size="24" fill="#0f172a">Service Health</text>
  <rect x="24" y="70" width="352" height="54" rx="8" fill="#dcfce7"/>
  <text x="42" y="103" font-size="18" fill="#166534">API  Healthy</text>
  <rect x="24" y="138" width="352" height="54" rx="8" fill="#fef3c7"/>
  <text x="42" y="171" font-size="18" fill="#92400e">Worker  Warning</text>
</svg>
'@
        [IO.File]::WriteAllText(
            $script:SvgPath,
            $svg,
            [Text.UTF8Encoding]::new($false))

        function New-TestTopologyInterchangeJson {
            param(
                [Parameter(Mandatory)]
                [string] $Id,
                [switch] $IncludeDatabase,
                [switch] $IncludeOwner
            )

            $apiExtensions = if ($IncludeOwner) { [ordered] @{ Owner = 'Platform' } } else { [ordered] @{} }
            $nodes = @(
                [ordered] @{
                    id         = if ($IncludeDatabase) { 'api' } else { 'service' }
                    role       = 'TopologyNode'
                    kind       = if ($IncludeDatabase) { 'Service' } else { 'Process' }
                    label      = if ($IncludeDatabase) { 'API' } else { 'Service' }
                    topology   = [ordered] @{
                        kind        = if ($IncludeDatabase) { 'Service' } else { 'Process' }
                        status      = 'Unknown'
                        displayMode = 'Card'
                    }
                    extensions = $apiExtensions
                    metrics    = @()
                    ports      = @()
                    details    = @()
                }
            )
            $edges = @()
            if ($IncludeDatabase) {
                $nodes += [ordered] @{
                    id         = 'database'
                    role       = 'TopologyNode'
                    kind       = 'Database'
                    label      = 'Database'
                    topology   = [ordered] @{
                        kind        = 'Database'
                        status      = 'Unknown'
                        displayMode = 'Card'
                    }
                    extensions = [ordered] @{}
                    metrics    = @()
                    ports      = @()
                    details    = @()
                }
                $edges = @(
                    [ordered] @{
                        id         = 'api-db'
                        role       = 'TopologyEdge'
                        kind       = 'DataFlow'
                        sourceId   = 'api'
                        targetId   = 'database'
                        label      = 'queries'
                        order      = 0
                        topology   = [ordered] @{
                            kind             = 'DataFlow'
                            status           = 'Unknown'
                            direction        = 'Forward'
                            sourcePort       = 'Auto'
                            targetPort       = 'Auto'
                            lineStyle        = 'Solid'
                            routing          = 'Orthogonal'
                            emphasis         = 'Normal'
                            layoutInference  = 'None'
                        }
                        extensions = [ordered] @{}
                        metrics    = @()
                    }
                )
            }

            [ordered] @{
                schema         = 'chartforgex.visual-artifact'
                version        = 1
                kind           = 'Topology'
                family         = 'Topology'
                sourceLanguage = 'Native'
                id             = $Id
                title          = if ($IncludeDatabase) { 'Portable topology' } else { '' }
                subtitle       = ''
                decorative     = $false
                topology       = [ordered] @{
                    layoutMode      = 'Layered'
                    layoutDirection = 'LeftToRight'
                }
                extensions     = [ordered] @{}
                groups         = @()
                nodes          = $nodes
                edges          = $edges
                scenarios      = @()
                annotations    = @()
            } | ConvertTo-Json -Depth 12 -Compress
        }
    }

    It 'converts once and reports the selected placement payload' {
        $visual = $SvgPath | ConvertTo-OfficeVisual -Width 300 -SvgPolicy RasterizeWhenNeeded -Id service-health -Title 'Service Health' -AlternativeText 'Health of API and worker services.'

        $visual.GetType().FullName | Should -Be 'OfficeIMO.ChartForgeX.OfficeVisualConversionResult'
        $visual.WidthPoints | Should -Be 300
        $visual.GetPlacementBytes().Count | Should -BeGreaterThan 100
        $visual.PlacementMediaType | Should -BeIn 'image/svg+xml', 'image/png'
        $visual.AlternativeText | Should -Be 'Health of API and worker services.'
    }

    It 'accepts a reusable Office visual source without treating it as a path' {
        $probe = $SvgPath | ConvertTo-OfficeVisual
        $sourceType = $probe.GetType().Assembly.GetType('OfficeIMO.ChartForgeX.OfficeVisualSource', $true)
        $source = [Activator]::CreateInstance($sourceType, (, [IO.File]::ReadAllBytes($SvgPath)))
        $source.Id = 'portable-health'
        $source.Title = 'Portable Health'
        $visual = $source | ConvertTo-OfficeVisual -Width 280

        $visual.Id | Should -Be 'portable-health'
        $visual.Title | Should -Be 'Portable Health'
        $visual.WidthPoints | Should -Be 280
        { $source | ConvertTo-OfficeVisual -Id changed -ErrorAction Stop } | Should -Throw '*cannot be used with an existing OfficeVisualSource*'
    }

    It 'keeps safe SVG viewport limits unless trusted input opts in explicitly' {
        $oversizedPath = Join-Path $TestDrive 'oversized.svg'
        [IO.File]::WriteAllText(
            $oversizedPath,
            '<svg xmlns="http://www.w3.org/2000/svg" width="9000" height="1" viewBox="0 0 9000 1"><rect width="9000" height="1" fill="#2563eb"/></svg>',
            [Text.UTF8Encoding]::new($false))

        { $oversizedPath | ConvertTo-OfficeVisual -ErrorAction Stop } | Should -Throw '*exceeds the configured import limits*'
        $trusted = $oversizedPath | ConvertTo-OfficeVisual -MaximumSvgViewportDimension 9000
        $trusted.WidthPoints | Should -Be 6750
        $trusted.HeightPoints | Should -Be 0.75
    }

    It 'accepts PathInfo and the ImagePlayground portable visual envelope' {
        $pathInfo = Resolve-Path -LiteralPath $SvgPath
        $fromPathInfo = $pathInfo | ConvertTo-OfficeVisual
        $fromPathInfo.GetSvgBytes().Count | Should -BeGreaterThan 100

        $driveName = 'VisualArtifacts' + [Guid]::NewGuid().ToString('N')
        try {
            New-PSDrive -Name $driveName -PSProvider FileSystem -Root $TestDrive | Out-Null
            $drivePathInfo = Resolve-Path -LiteralPath "${driveName}:\service-health.svg"
            $fromDrivePathInfo = $drivePathInfo | ConvertTo-OfficeVisual
            $fromDrivePathInfo.GetSvgBytes().Count | Should -BeGreaterThan 100
        } finally {
            Remove-PSDrive -Name $driveName -ErrorAction SilentlyContinue
        }

        $portable = [pscustomobject] @{
            OfficeVisualSvg             = [IO.File]::ReadAllBytes($SvgPath)
            OfficeVisualId              = 'portable-pipeline'
            OfficeVisualTitle           = 'Portable pipeline'
            OfficeVisualAlternativeText = 'Portable visual across module load contexts.'
        }
        $portable.PSObject.TypeNames.Insert(0, 'ImagePlayground.VisualArtifact')
        $visual = $portable | ConvertTo-OfficeVisual -Width 260

        $visual.Id | Should -Be 'portable-pipeline'
        $visual.Title | Should -Be 'Portable pipeline'
        $visual.AlternativeText | Should -Be 'Portable visual across module load contexts.'
        $visual.WidthPoints | Should -Be 260
    }

    It 'converts portable semantic bytes into native editable Visio objects' {
        $portable = [pscustomobject] @{
            OfficeVisualInterchangeJson = [Text.Encoding]::UTF8.GetBytes((New-TestTopologyInterchangeJson -Id portable-topology -IncludeDatabase -IncludeOwner))
            OfficeVisualSvg             = [IO.File]::ReadAllBytes($SvgPath)
        }
        $portable.PSObject.TypeNames.Insert(0, 'ImagePlayground.VisualArtifact')

        $converted = $portable | ConvertTo-OfficeVisioVisual -PageName 'Portable topology'
        $converted.GetType().FullName | Should -Be 'OfficeIMO.ChartForgeX.OfficeVisioVisualConversionResult'
        $converted.Report.AllProjectedObjectsEditable | Should -BeTrue
        $converted.Report.NodeCount | Should -Be 2
        $converted.Page.Shapes.Id | Should -Contain 'api'
        $converted.Page.Shapes.Id | Should -Contain 'database'
        $converted.Page.Connectors.Id | Should -Contain 'api-db'
        ($converted.Page.Shapes | Where-Object Id -EQ 'api').GetShapeDataValue('Extension.Owner') | Should -Be 'Platform'

        $natural = $portable | ConvertTo-OfficeVisioVisual -UseNaturalPageSize -PixelsPerInch 100
        $natural.Page.Width | Should -BeGreaterThan $converted.Page.Width
    }

    It 'exports portable semantic bytes as a valid editable VSDX package' {
        $portable = [pscustomobject] @{
            OfficeVisualInterchangeJson = [Text.Encoding]::UTF8.GetBytes((New-TestTopologyInterchangeJson -Id export-topology))
        }
        $portable.PSObject.TypeNames.Insert(0, 'ImagePlayground.VisualArtifact')
        $path = Join-Path $TestDrive 'portable-topology.vsdx'

        $file = $portable | Export-OfficeVisioVisual -Path $path

        $file.FullName | Should -Be $path
        Test-Path $path | Should -BeTrue
        $loaded = Get-OfficeVisio -Path $path
        $loaded.Pages[0].Shapes.Id | Should -Contain 'service'
    }

    It 'accepts raw semantic JSON text without interpreting it as a file path' {
        $json = New-TestTopologyInterchangeJson -Id raw-json

        $converted = $json | ConvertTo-OfficeVisioVisual

        $converted.Envelope.Id | Should -Be 'raw-json'
        $converted.Page.Shapes.Id | Should -Contain 'service'
    }

    It 'accepts raw semantic JSON text with a leading Unicode BOM' {
        $json = [char] 0xFEFF + (New-TestTopologyInterchangeJson -Id bom-json)

        $converted = $json | ConvertTo-OfficeVisioVisual

        $converted.Envelope.Id | Should -Be 'bom-json'
        $converted.Page.Shapes.Id | Should -Contain 'service'
    }

    It 'accepts semantic JSON PathInfo from a custom filesystem PSDrive' {
        $jsonPath = Join-Path $TestDrive 'custom-drive.json'
        [IO.File]::WriteAllText(
            $jsonPath,
            (New-TestTopologyInterchangeJson -Id custom-drive),
            [Text.UTF8Encoding]::new($false))
        $driveName = 'VisualInterchange' + [Guid]::NewGuid().ToString('N')

        try {
            New-PSDrive -Name $driveName -PSProvider FileSystem -Root $TestDrive | Out-Null
            $pathInfo = Resolve-Path -LiteralPath "${driveName}:\custom-drive.json"

            $converted = $pathInfo | ConvertTo-OfficeVisioVisual

            $converted.Envelope.Id | Should -Be 'custom-drive'
            $converted.Page.Shapes.Id | Should -Contain 'service'
        } finally {
            Remove-PSDrive -Name $driveName -ErrorAction SilentlyContinue
        }
    }

    It 'reassembles normally piped semantic JSON bytes for conversion and export' {
        [byte[]] $bytes = [Text.Encoding]::UTF8.GetBytes((New-TestTopologyInterchangeJson -Id piped-bytes))

        $converted = $bytes | ConvertTo-OfficeVisioVisual
        $converted.Envelope.Id | Should -Be 'piped-bytes'
        $converted.Page.Shapes.Id | Should -Contain 'service'

        $path = Join-Path $TestDrive 'piped-bytes.vsdx'
        $file = $bytes | Export-OfficeVisioVisual -Path $path
        $file.FullName | Should -Be $path
        (Get-OfficeVisio -Path $path).Pages[0].Shapes.Id | Should -Contain 'service'
    }

    It 'rejects mixed byte-stream and typed pipeline input' {
        [byte[]] $bytes = [Text.Encoding]::UTF8.GetBytes((New-TestTopologyInterchangeJson -Id mixed-input))

        { @($bytes[0], (New-TestTopologyInterchangeJson -Id second-input)) |
                ConvertTo-OfficeVisioVisual -ErrorAction Stop } |
            Should -Throw '*cannot be mixed*'
    }

    It 'rejects multiple pipeline artifacts for one VSDX destination' {
        $portable = [pscustomobject] @{
            OfficeVisualInterchangeJson = [Text.Encoding]::UTF8.GetBytes((New-TestTopologyInterchangeJson -Id single-output))
        }
        $portable.PSObject.TypeNames.Insert(0, 'ImagePlayground.VisualArtifact')
        $path = Join-Path $TestDrive 'single-output.vsdx'

        { @($portable, $portable) | Export-OfficeVisioVisual -Path $path -ErrorAction Stop } |
            Should -Throw '*accepts one input artifact*'
        Test-Path $path | Should -BeFalse
    }

    It 'rejects oversized semantic JSON files before reading their payload' {
        $path = Join-Path $TestDrive 'oversized.json'
        $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
        try {
            $stream.SetLength((8MB) + 1)
        } finally {
            $stream.Dispose()
        }

        { ConvertTo-OfficeVisioVisual -InputObject $path -ErrorAction Stop } |
            Should -Throw '*must not exceed*bytes*'
    }

    It 'rejects an SVG-only ImagePlayground envelope for editable Visio conversion' {
        $portable = [pscustomobject] @{ OfficeVisualSvg = [IO.File]::ReadAllBytes($SvgPath) }
        $portable.PSObject.TypeNames.Insert(0, 'ImagePlayground.VisualArtifact')

        { $portable | ConvertTo-OfficeVisioVisual -ErrorAction Stop } |
            Should -Throw '*OfficeVisualInterchangeJson*'
    }

    It 'places one converted visual in Word, Excel, and PowerPoint' {
        $visual = $SvgPath | ConvertTo-OfficeVisual -Width 300 -AlternativeText 'Health of API and worker services.'
        $wordPath = Join-Path $TestDrive 'visual.docx'
        $excelPath = Join-Path $TestDrive 'visual.xlsx'
        $powerPointPath = Join-Path $TestDrive 'visual.pptx'

        New-OfficeWord -Path $wordPath {
            WordSection { WordParagraph { $visual | Add-OfficeWordVisual | Out-Null } }
        } | Out-Null
        New-OfficeExcel -Path $excelPath {
            Add-OfficeExcelSheet -Name Dashboard -Content {
                $visual | Add-OfficeExcelVisual -Address B2 | Out-Null
            }
        } | Out-Null
        New-OfficePowerPoint -Path $powerPointPath {
            PptSlide { $visual | Add-OfficePowerPointVisual -X 36 -Y 54 | Out-Null }
        } | Out-Null
        Test-Path $wordPath | Should -BeTrue
        Test-Path $excelPath | Should -BeTrue
        Test-Path $powerPointPath | Should -BeTrue
    }

    It 'places one converted visual through both PDF composition paths' {
        $visual = $SvgPath | ConvertTo-OfficeVisual -Width 300 -AlternativeText 'Health of API and worker services.'
        $dslPath = Join-Path $TestDrive 'visual-dsl.pdf'
        $documentPath = Join-Path $TestDrive 'visual-document.pdf'

        New-OfficePdf -Path $dslPath {
            $visual | Add-OfficePdfVisual -Align Center
        } | Out-Null

        $document = New-OfficePdf -Content { PdfParagraph 'Existing document' }
        $updated = $visual | Add-OfficePdfVisual -Document $document -PassThru
        [object]::ReferenceEquals($updated, $document) | Should -BeTrue
        $pipelineUpdated = $document | Add-OfficePdfVisual -InputObject $visual -PassThru
        [object]::ReferenceEquals($pipelineUpdated, $document) | Should -BeTrue
        $document | Save-OfficePdf -Path $documentPath | Out-Null

        Test-Path $dslPath | Should -BeTrue
        Test-Path $documentPath | Should -BeTrue
    }

    It 'rejects conversion overrides when a prepared visual is reused' {
        $visual = $SvgPath | ConvertTo-OfficeVisual -Width 300
        $wordPath = Join-Path $TestDrive 'override.docx'
        $document = New-OfficeWord -Path $wordPath -NoSave
        try {
            $paragraph = $document.AddParagraph()
            { $visual | Add-OfficeWordVisual -Paragraph $paragraph -Width 240 -ErrorAction Stop } | Should -Throw '*cannot be used with an existing*'
        } finally {
            Close-OfficeWord -Document $document -ErrorAction SilentlyContinue
        }
    }
}
