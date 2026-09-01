$wslExecutable = Get-BenchmarkInput WslExecutable 'wsl.exe'
$runnerPath = Get-BenchmarkInput RunnerPath -Required
$fixtureRoot = Get-BenchmarkInput FixtureRoot -Required
$tesseractRoot = Get-BenchmarkInput TesseractRoot -Required
$rapidPackages = Get-BenchmarkInput RapidPackages -Required
$rapidModels = Get-BenchmarkInput RapidModels -Required
$manifestPath = Get-BenchmarkInput ManifestPath -Required
$tesseractVersion = Get-BenchmarkInput TesseractVersion -Required
$rapidOcrVersion = Get-BenchmarkInput RapidOcrVersion -Required
$onnxRuntimeVersion = Get-BenchmarkInput OnnxRuntimeVersion -Required
$tesseractFootprintBytes = Get-BenchmarkInput TesseractFootprintBytes -Required
$rapidFootprintBytes = Get-BenchmarkInput RapidFootprintBytes -Required

$cases = @(Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json)

function Get-OcrNormalizedText {
    param([AllowEmptyString()][string] $Text)

    $builder = [Text.StringBuilder]::new()
    foreach ($character in $Text.Normalize([Text.NormalizationForm]::FormC).ToLowerInvariant().ToCharArray()) {
        if ([char]::IsLetterOrDigit($character)) {
            [void] $builder.Append($character)
        }
    }
    $builder.ToString()
}

function Get-OcrCharacterErrorRate {
    param(
        [AllowEmptyString()][string] $Expected,
        [AllowEmptyString()][string] $Actual
    )

    $left = Get-OcrNormalizedText $Expected
    $right = Get-OcrNormalizedText $Actual
    if ($left.Length -eq 0) {
        return $right.Length -eq 0 ? 0.0 : 1.0
    }
    $previous = [int[]]::new($right.Length + 1)
    $current = [int[]]::new($right.Length + 1)
    for ($column = 0; $column -le $right.Length; $column++) {
        $previous[$column] = $column
    }
    for ($row = 1; $row -le $left.Length; $row++) {
        $current[0] = $row
        for ($column = 1; $column -le $right.Length; $column++) {
            $cost = $left[$row - 1] -eq $right[$column - 1] ? 0 : 1
            $current[$column] = [Math]::Min(
                [Math]::Min($current[$column - 1] + 1, $previous[$column] + 1),
                $previous[$column - 1] + $cost)
        }
        $swap = $previous
        $previous = $current
        $current = $swap
    }
    $previous[$right.Length] / [double] $left.Length
}

function Invoke-OcrComparisonLane {
    param(
        [Parameter(Mandatory)][ValidateSet('tesseract', 'rapidocr')][string] $Engine,
        [Parameter(Mandatory)] $Case,
        [Parameter(Mandatory)] $Run
    )

    $imagePath = $fixtureRoot.TrimEnd('/') + '/' + $Case.Image
    $output = @(& $wslExecutable -- env "PYTHONDONTWRITEBYTECODE=1" python3 $runnerPath `
        --engine $Engine `
        --image $imagePath `
        --language $Case.Language `
        --tesseract-root $tesseractRoot `
        --rapid-packages $rapidPackages `
        --rapid-models $rapidModels 2>&1)
    if ($LASTEXITCODE -ne 0) {
        throw "$Engine failed for $($Case.Name): $($output -join [Environment]::NewLine)"
    }
    $json = @($output | Where-Object { $_ -is [string] -and $_.TrimStart().StartsWith('{') } | Select-Object -Last 1)
    if ($json.Count -ne 1) {
        throw "$Engine did not emit one JSON result for $($Case.Name)."
    }
    $Run.Recognition = $json[0] | ConvertFrom-Json
    $Run.CharacterErrorRate = Get-OcrCharacterErrorRate -Expected $Case.Expected -Actual $Run.Recognition.text
    $footprintBytes = $Engine -eq 'tesseract' ? [long] $tesseractFootprintBytes : [long] $rapidFootprintBytes
    $Run.FootprintMiB = [Math]::Round($footprintBytes / 1MB, 2)
}

New-BenchmarkSuite 'officeimo-ocr-engine-comparison' {
    Set-BenchmarkPolicy -Warmup 0 -Iterations 3 -Order Rotated -MemoryCleanup None -OutlierMode None
    Set-BenchmarkProfile Current -Cleanup KeepOnFailure

    Add-BenchmarkCaseSource {
        foreach ($item in $cases) {
            [pscustomobject]@{
                Name = $item.name
                Image = $item.image
                Language = $item.language
                Expected = $item.expected
            }
        }
    }

    Add-BenchmarkEngine Tesseract {
        Add-BenchmarkOperation Recognize {
            param($case, $run)
            Invoke-OcrComparisonLane -Engine tesseract -Case $case -Run $run
        }
    }

    Add-BenchmarkEngine RapidOCR {
        Add-BenchmarkOperation Recognize {
            param($case, $run)
            Invoke-OcrComparisonLane -Engine rapidocr -Case $case -Run $run
        }
    }

    Add-BenchmarkValidation {
        param($case, $run)
        if ($null -eq $run.Recognition -or [string]::IsNullOrWhiteSpace($run.Recognition.text)) {
            throw "$($case.Name) produced no recognized text."
        }
        if (@($run.Recognition.words).Count -eq 0) {
            throw "$($case.Name) produced no recognized geometry."
        }
        foreach ($word in @($run.Recognition.words)) {
            if ($word.width -le 0 -or $word.height -le 0) {
                throw "$($case.Name) produced non-positive OCR geometry."
            }
        }
        if ($run.CharacterErrorRate -gt 0.5) {
            throw "$($case.Name) exceeded the 50 percent character-error validity ceiling."
        }
    }

    Add-BenchmarkMetric CharacterErrorRate {
        param($case, $run)
        $run.CharacterErrorRate
    }
    Add-BenchmarkMetric RecognizedRegions {
        param($case, $run)
        @($run.Recognition.words).Count
    }
    Add-BenchmarkMetric MeanConfidence {
        param($case, $run)
        (@($run.Recognition.words) | Measure-Object confidence -Average).Average
    }
    Add-BenchmarkMetric InstalledMiB {
        param($case, $run)
        $run.FootprintMiB
    }

    Add-BenchmarkComparison Engine -Baseline Tesseract -Metric MedianMs -TieTolerance 0.05
    Add-BenchmarkMetadata ComparisonMode 'Equivalent one-shot local CPU OCR through WSL process boundaries'
    Add-BenchmarkMetadata TesseractVersion $tesseractVersion
    Add-BenchmarkMetadata RapidOcrVersion $rapidOcrVersion
    Add-BenchmarkMetadata OnnxRuntimeVersion $onnxRuntimeVersion
    Set-BenchmarkArtifacts Json, Csv, Markdown
}
