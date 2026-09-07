function Invoke-ShowcaseDotNet {
    param([Parameter(Mandatory)][string[]] $Arguments)
    & dotnet @Arguments
    if ($LASTEXITCODE -ne 0) { throw "dotnet failed with exit code $LASTEXITCODE." }
}

function Resolve-ShowcasePath {
    param([string] $Root, [string] $RelativePath)
    $absoluteRoot = [System.IO.Path]::GetFullPath($Root)
    $target = [System.IO.Path]::GetFullPath((Join-Path $absoluteRoot $RelativePath))
    $comparison = if ($IsWindows) { [StringComparison]::OrdinalIgnoreCase } else { [StringComparison]::Ordinal }
    if (-not $target.StartsWith($absoluteRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar, $comparison)) {
        throw "Showcase path escapes its declared root: $RelativePath"
    }
    return $target
}

function New-ShowcasePdfPreview {
    param([string] $InputPath, [string] $OutputPath, [int] $Page = 1)
    $renderer = Get-Command 'pdftocairo' -ErrorAction Stop
    if ($Page -lt 1) { throw 'PDF preview pages are one-based.' }
    New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
    $outputBase = Join-Path (Split-Path -Parent $OutputPath) ([IO.Path]::GetFileNameWithoutExtension($OutputPath))
    & $renderer.Source '-f' $Page '-l' $Page '-singlefile' '-png' '-scale-to' '1123' $InputPath $outputBase
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
        throw "Could not render showcase PDF: $InputPath"
    }
}

function New-ShowcaseReaderProjection {
    param([string] $InputPath, [string] $OutputPath, [string] $RepositoryRoot, [string] $Configuration, [string] $Framework)
    Invoke-ShowcaseDotNet @('build', (Join-Path $RepositoryRoot 'OfficeIMO.Tool/OfficeIMO.Tool.csproj'), '-c', $Configuration, '-f', $Framework, '--nologo')
    $readerAssembly = Join-Path $RepositoryRoot "OfficeIMO.Tool/bin/$Configuration/$Framework/OfficeIMO.Tool.dll"
    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'dotnet'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    foreach ($argument in @($readerAssembly, 'reader', 'read', '-', '--name', 'design-brief.pptx', '--format', 'json', '--output', $OutputPath)) {
        [void] $startInfo.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::Start($startInfo)
    try {
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $bytes = [IO.File]::ReadAllBytes($InputPath)
        $process.StandardInput.BaseStream.Write($bytes, 0, $bytes.Length)
        $process.StandardInput.Close()
        $process.WaitForExit()
        [void] $stdout.GetAwaiter().GetResult()
        $errorText = $stderr.GetAwaiter().GetResult()
        if ($process.ExitCode -ne 0) { throw "Reader projection failed: $errorText" }
    } finally {
        $process.Dispose()
    }
}
