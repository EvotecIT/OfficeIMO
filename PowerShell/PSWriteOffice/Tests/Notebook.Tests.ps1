BeforeAll {
    $moduleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        $env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES = 'true'
        if (-not $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION) {
            $release = Join-Path $PSScriptRoot '../Sources/PSWriteOffice/bin/Release'
            $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = if (Test-Path $release) { 'Release' } else { 'Debug' }
        }
        Join-Path $PSScriptRoot '../PSWriteOffice.psd1'
    }
    Import-Module $moduleManifest -Global -ErrorAction Stop
}

Describe 'Notebook command workflows' {
    It 'opens a literal FileInfo from the pipeline and saves in-memory edits through the associated path' {
        $path = Join-Path $TestDrive 'Lab[review].ipynb'
        $notebook = New-OfficeNotebook -Path $path
        try {
            Test-Path -LiteralPath $path | Should -BeFalse
            $notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Review notes'
            $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'Get-Date' -PassThru
            $cell | Set-OfficeNotebookCell -ExecutionCount 7
            $cell | Add-OfficeNotebookOutput -Text 'recorded result'
            $notebook | Add-OfficeNotebookCell -Kind Raw -Source 'scratch'
            $notebook | Move-OfficeNotebookCell -Index 1 -DestinationIndex 0
            $notebook | Remove-OfficeNotebookCell -Index 2
            $notebook | Remove-OfficeNotebookCell -Index 0 -WhatIf
            $notebook | Save-OfficeNotebook -WhatIf
            Test-Path -LiteralPath $path | Should -BeFalse
            ($notebook | Save-OfficeNotebook -PassThru).FullName | Should -Be $path
        } finally {
            $notebook | Close-OfficeNotebook
        }

        $loaded = Get-Item -LiteralPath $path | Open-OfficeNotebook
        try {
            @($loaded | Get-OfficeNotebookCell).Count | Should -Be 2
            $code = $loaded | Get-OfficeNotebookCell -Index 0 -Kind Code
            $code.Language | Should -Be 'PowerShell'
            $code.ExecutionCount | Should -Be 7
            ($code | Get-OfficeNotebookOutput -Index 0).Text | Should -Be 'recorded result'
            $code | Set-OfficeNotebookCell -Source 'Get-Location' -ExecutionCount $null
            @($code | Get-OfficeNotebookOutput).Count | Should -Be 1
            $code | Clear-OfficeNotebookOutput -WhatIf
            @($code | Get-OfficeNotebookOutput).Count | Should -Be 1
            $code | Clear-OfficeNotebookOutput
            $loaded | Save-OfficeNotebook
        } finally {
            $loaded | Close-OfficeNotebook
        }

        $saved = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        $saved.cells[0].source | Should -Be 'Get-Location'
        $saved.cells[0].execution_count | Should -BeNullOrEmpty
        $saved.cells[0].outputs | Should -BeNullOrEmpty
        $saved.cells[1].source | Should -Be '# Review notes'
    }

    It 'matches a caller-selected companion through the explicit Jupytext profile and reports a mismatched file' {
        $source = Join-Path $TestDrive 'paired.py'
        $companion = Join-Path $TestDrive 'saved.ipynb'
        [IO.File]::WriteAllText($source, "# %%`nprint(42)`n")
        [IO.File]::WriteAllText($companion, @'
{"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[{"id":"result","cell_type":"code","metadata":{},"source":"print(42)","execution_count":2,"outputs":[{"output_type":"stream","name":"stdout","text":"saved forty-two"}]}]}
'@)

        { Open-OfficeNotebook -LiteralPath $source -ErrorAction Stop } | Should -Throw '*Specify -Format*'
        $notebook = Open-OfficeNotebook -LiteralPath $source -Format JupytextPercent -CompanionPath $companion -FailOnLoss
        try {
            ($notebook | Get-OfficeNotebookCell -Index 0 | Get-OfficeNotebookOutput -Index 0).Text |
                Should -Be 'saved forty-two'
            { $notebook | Save-OfficeNotebook -ErrorAction Stop } | Should -Throw '*Supply -Path*'
        } finally {
            $notebook | Close-OfficeNotebook
        }

        [IO.File]::WriteAllText($source, "# %%`nprint(43)`n")
        { Open-OfficeNotebook -LiteralPath $source -Format JupytextPercent -CompanionPath $companion -FailOnLoss -ErrorAction Stop } |
            Should -Throw -ErrorId 'notebook-companion-mismatch,*'
    }

    It 'pipes one saved structured data resource to the CSV and Excel file adapters' {
        $source = Join-Path $TestDrive 'table.ipynb'
        $csv = Join-Path $TestDrive 'table.csv'
        $xlsx = Join-Path $TestDrive 'table.xlsx'
        [IO.File]::WriteAllText($source, @'
{"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[{"id":"data","cell_type":"code","metadata":{},"source":"data","execution_count":1,"outputs":[{"output_type":"display_data","metadata":{},"data":{"application/vnd.dataresource+json":{"schema":{"fields":[{"name":"City","type":"string"},{"name":"Count","type":"integer"}]},"data":[{"City":"Krakow","Count":42}]}}}]}]}
'@)

        $notebook = Open-OfficeNotebook -LiteralPath $source
        try {
            $result = $notebook | Get-OfficeNotebookCell -Index 0 | Get-OfficeNotebookOutput -Index 0
            ($result | Export-OfficeNotebook -Path $csv).OutputPath | Should -Be $csv
            ($result | Export-OfficeNotebook -Path $xlsx).ItemsWritten | Should -Be 1
        } finally {
            $notebook | Close-OfficeNotebook
        }

        $csvRow = Import-Csv -LiteralPath $csv
        $csvRow.City | Should -Be 'Krakow'
        $csvRow.Count | Should -Be '42'
        $sheetRow = Import-OfficeExcel -Path $xlsx -Sheet 'Notebook data'
        $sheetRow.City | Should -Be 'Krakow'
        $sheetRow.Count | Should -Be 42
    }

    It 'exports a live document and rejects no-clobber without changing the original output' {
        $path = Join-Path $TestDrive 'report.html'
        $notebook = New-OfficeNotebook
        try {
            $notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Review report'
            $notebook | Add-OfficeNotebookCell -Kind Code -Source 'hidden-code-marker'
            ($notebook | Export-OfficeNotebook -Path $path -IncludeCode:$false).OutputPath | Should -Be $path
            $report = Get-Content -LiteralPath $path -Raw
            $report | Should -Match '<h1[^>]*>Review report</h1>'
            $report | Should -Not -Match 'hidden-code-marker'
            { $notebook | Export-OfficeNotebook -Path $path -NoClobber -ErrorAction Stop } | Should -Throw '*already exists*'
            (Get-Content -LiteralPath $path -Raw) | Should -BeExactly $report
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }

    It 'can preserve Markdown cell spelling in a text export' {
        $path = Join-Path $TestDrive 'source.md'
        $notebook = New-OfficeNotebook
        try {
            $source = "Source heading`r`n=============="
            $notebook | Add-OfficeNotebookCell -Kind Markdown -Source $source
            $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Source 'example()' -PassThru
            $cell | Add-OfficeNotebookOutput -Text 'saved result'
            $notebook | Export-OfficeNotebook -Path $path -PreserveMarkdownSource | Out-Null
            $text = [IO.File]::ReadAllText($path)
            $text.StartsWith($source, [StringComparison]::Ordinal) | Should -BeTrue
            $text | Should -Match 'saved result'
            { $notebook | Export-OfficeNotebook -Path (Join-Path $TestDrive 'source.html') -PreserveMarkdownSource -ErrorAction Stop } |
                Should -Throw '*only to Markdown exports*'
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }

    It 'can opt in to saved controls with the HTML export switch' {
        $sample = Join-Path $PSScriptRoot '../../../OfficeIMO.TestAssets/Notebook/AdvancedOutputs/Widgets/widget-list.ipynb'
        $path = Join-Path $TestDrive 'widgets.html'
        $notebook = Open-OfficeNotebook -LiteralPath $sample
        try {
            $notebook | Export-OfficeNotebook -Path $path -IncludeInteractiveWidgets -IncludeCode:$false | Out-Null
            [IO.File]::ReadAllText($path) | Should -Match 'data-officeimo-widget-kind="play"'
            { $notebook | Export-OfficeNotebook -Path (Join-Path $TestDrive 'widgets.pdf') -IncludeInteractiveWidgets -ErrorAction Stop } |
                Should -Throw '*only to HTML exports*'
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }

    It 'round trips static PowerShell source and rejects saved output loss in strict mode' {
        $script = Join-Path $TestDrive 'work.ps1'
        $marker = Join-Path $TestDrive 'execution-witness.txt'
        $notebook = New-OfficeNotebook
        $source = "[IO.File]::WriteAllText('$($marker.Replace("'", "''"))', 'Executed')"
        try {
            $notebook | Add-OfficeNotebookCell -Kind Markdown -Source '# Lab notes'
            $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source $source -PassThru
            $cell | Add-OfficeNotebookOutput -Text 'stored output'
            [IO.File]::WriteAllText($script, 'original')
            { $notebook | Export-OfficeNotebook -Path $script -FailOnLoss -ErrorAction Stop } | Should -Throw
            (Get-Content -LiteralPath $script -Raw) | Should -BeExactly 'original'
            ($notebook | Export-OfficeNotebook -Path $script).HasLoss | Should -BeTrue
        } finally {
            $notebook | Close-OfficeNotebook
        }

        $loaded = Open-OfficeNotebook -LiteralPath $script -Format PowerShellScript
        try {
            ($loaded | Get-OfficeNotebookCell -Index 0).Source | Should -BeExactly '# Lab notes'
            ($loaded | Get-OfficeNotebookCell -Index 1).Source | Should -BeExactly $source
            Test-Path -LiteralPath $marker | Should -BeFalse
        } finally {
            $loaded | Close-OfficeNotebook
        }
    }

    It 'runs a literal notebook in document order and saves the selected outputs' {
        $source = Join-Path $TestDrive 'Lab[execution].ipynb'
        $destination = Join-Path $TestDrive 'executed.ipynb'
        $notebook = New-OfficeNotebook -Path $source
        try {
            $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source '$shared = 41'
            $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source 'throw "This cell was not selected"'
            $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source '$shared + 1; "$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"'
            $notebook | Save-OfficeNotebook
        } finally {
            $notebook | Close-OfficeNotebook
        }
        $inputHash = (Get-FileHash -LiteralPath $source).Hash

        $items = @(Get-Item -LiteralPath $source |
            Invoke-OfficeNotebook -CellIndex 2, 0 -StreamOutput -OutputPath $destination -NoClobber -ErrorAction Stop)
        $report = $items[-1]
        try {
            $items[0].GetType().Name | Should -Be 'NotebookExecutionEvent'
            $report.Cells.Count | Should -Be 2
            $saved = Get-Content -LiteralPath $destination -Raw | ConvertFrom-Json
            $saved.cells[0].execution_count | Should -Be 1
            $saved.cells[1].execution_count | Should -BeNullOrEmpty
            $saved.cells[2].execution_count | Should -Be 2
            $saved.cells[2].outputs[0].output_type | Should -Be 'execute_result'
            $saved.cells[2].outputs[0].data.'text/plain' | Should -BeExactly '42'
            $workerVersion = $saved.cells[2].outputs[1].data.'text/plain'
            $workerVersion | Should -BeExactly "$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)"
            Write-Host "Notebook worker verified as $workerVersion"
            (Get-FileHash -LiteralPath $source).Hash | Should -BeExactly $inputHash
        } finally {
            $report.Notebook | Close-OfficeNotebook
        }
    }

    It 'runs cells in this process unless a separate process is requested' {
        $notebook = New-OfficeNotebook
        try {
            $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source '$PID' -PassThru
            $null = $notebook | Invoke-OfficeNotebook -ErrorAction Stop
            ($cell | Get-OfficeNotebookOutput -Index 0).Data['text/plain'].GetString() | Should -BeExactly "$PID"

            $null = $notebook | Invoke-OfficeNotebook -Isolation Process -ErrorAction Stop
            ($cell | Get-OfficeNotebookOutput -Index 0).Data['text/plain'].GetString() | Should -Not -Be "$PID"

            { $notebook | Invoke-OfficeNotebook -Isolation Runspace -PowerShellPath (Get-Process -Id $PID).Path -ErrorAction Stop } |
                Should -Throw '*separate process*'
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }

    It 'protects live notebooks and destinations before executing trusted code' {
        $destination = Join-Path $TestDrive 'existing.ipynb'
        $marker = Join-Path $TestDrive 'invoked.txt'
        $notebook = New-OfficeNotebook
        $code = "[IO.File]::WriteAllText('$($marker.Replace("'", "''"))', 'ran'); 'fresh result'"
        [IO.File]::WriteAllText($destination, 'original destination')
        try {
            $cell = $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source $code -PassThru
            $cell | Add-OfficeNotebookOutput -Text 'stored result'
            $notebook | Invoke-OfficeNotebook -OutputPath $destination -WhatIf
            { $notebook | Invoke-OfficeNotebook -OutputPath $destination -NoClobber -ErrorAction Stop } |
                Should -Throw '*already exists*'
            Test-Path -LiteralPath $marker | Should -BeFalse
            [IO.File]::ReadAllText($destination) | Should -BeExactly 'original destination'
            ($cell | Get-OfficeNotebookOutput -Index 0).Text | Should -BeExactly 'stored result'

            $report = $notebook | Invoke-OfficeNotebook -ErrorAction Stop
            [object]::ReferenceEquals($notebook, $report.Notebook) | Should -BeTrue
            [IO.File]::ReadAllText($marker) | Should -BeExactly 'ran'
            ($cell | Get-OfficeNotebookOutput -Index 0).Data['text/plain'].GetString() | Should -BeExactly 'fresh result'
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }

    It 'saves captured output when a cell times out' {
        $destination = Join-Path $TestDrive 'timed-out.ipynb'
        $notebook = New-OfficeNotebook
        try {
            $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source '"before deadline"; Start-Sleep -Seconds 10'
            $notebook | Add-OfficeNotebookCell -Kind Code -Language PowerShell -Source '"later cell"'
            $report = $notebook | Invoke-OfficeNotebook -TimeoutSeconds 3 -ContinueOnError -OutputPath $destination -ErrorAction Stop

            $report.Cells.Count | Should -Be 1
            $report.Cells[0].Status.ToString() | Should -Be 'TimedOut'
            $saved = Get-Content -LiteralPath $destination -Raw | ConvertFrom-Json
            $saved.cells[0].outputs[0].data.'text/plain' | Should -BeExactly 'before deadline'
            $saved.cells[0].outputs[-1].ename | Should -Be 'NotebookCellTimeout'
            $saved.cells[1].execution_count | Should -BeNullOrEmpty
        } finally {
            $notebook | Close-OfficeNotebook
        }
    }
}
