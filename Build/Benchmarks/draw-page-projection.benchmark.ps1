$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$binaryRoot = Get-BenchmarkInput BinaryRoot (Join-Path $repositoryRoot 'OfficeIMO.OpenDocument.Benchmarks/bin/Release/net8.0')
$pageCounts = Get-BenchmarkInput PageCounts 100, 500, 1000 -Int
$includeEmpty = Get-BenchmarkInput IncludeEmpty $true -Bool
$includeFields = Get-BenchmarkInput IncludeFields $true -Bool
$workloads = @{}
[void] [Reflection.Assembly]::LoadFrom((Join-Path $binaryRoot 'OfficeIMO.OpenDocument.Benchmarks.dll'))

New-BenchmarkSuite 'officeimo-draw-page-projection' {
    Set-BenchmarkPolicy -Warmup 2 -Iteration 5 -Order Rotated -OutlierMode None -MemoryCleanup BeforeIteration
    Add-BenchmarkMetadata Contract 'Project every page using ToDrawings, with one shared master/layout. Validate every page dimension, visible page/count field and unchanged source XML outside measurement.'
    Add-BenchmarkMetadata Measurement 'Shared runner elapsed time and managed allocation include host invocation; input generation and readback are excluded. Working-set delta is neither retained nor peak memory.'
    Add-BenchmarkMetadata AffinityPolicy 'Inherited; macOS processor placement and busy-host timing are unqualified. Allocation evidence is host-specific.'
    Add-BenchmarkMetadata Runtime ([Runtime.InteropServices.RuntimeInformation]::FrameworkDescription)
    $runnerPath = Join-Path (Split-Path (Get-Command Invoke-BenchmarkSuite).ImplementingType.Assembly.Location) 'PowerForge.PowerShell.dll'
    Add-BenchmarkMetadata RunnerSha256 (Get-FileHash $runnerPath).Hash
    foreach ($name in 'OfficeIMO.Core', 'OfficeIMO.OpenDocument', 'OfficeIMO.OpenDocument.Benchmarks') {
        Add-BenchmarkMetadata "$name.Sha256" (Get-FileHash (Join-Path $binaryRoot "$name.dll")).Hash
    }
    Add-BenchmarkCaseSource {
        foreach ($pages in $pageCounts) {
            foreach ($kind in 'Empty', 'Fields') {
                if ($kind -eq 'Empty' -and -not $includeEmpty -or $kind -eq 'Fields' -and -not $includeFields) { continue }
                $fields = $kind -eq 'Fields'
                $key = "$kind-$pages"
                $workloads[$key] = [OfficeIMO.OpenDocument.Benchmarks.OdgProjectionWorkload]::new($pages, $fields)
                [pscustomobject]@{ Name = $key; Kind = $kind; Pages = $pages; Fields = $fields; InputSha256 = $workloads[$key].InputSha256 }
            }
        }
    }
    Set-BenchmarkSetup {
        param($case, $run)
        $run.Workload = $workloads["$($case.Kind)-$($case.Pages)"]
        $run.Workload.ReleaseResults()
    }
    Add-BenchmarkEngine OfficeIMO {
        Add-BenchmarkOperation ProjectPages {
            param($case, $run)
            $run.Workload.Execute()
        }
    }
    Add-BenchmarkValidation {
        param($case, $run)
        $run.Workload.Validate()
        Assert-BenchmarkValue -Actual $run.Workload.VerifiedPages -Expected $case.Pages
        $run.Workload.ReleaseResults()
    }
    Add-BenchmarkMetric VerifiedPages { param($case, $run) $run.Workload.VerifiedPages }
    Add-BenchmarkMetric ContentChecksum { param($case, $run) $run.Workload.ContentChecksum }
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
