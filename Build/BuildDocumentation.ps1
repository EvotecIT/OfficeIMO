param (
    $SolutionRoot = "$PSScriptRoot\.."
)

$DotNetVersion = 'net8.0'
$SolutionPath = [io.path]::Combine($SolutionRoot, 'OfficeImo.sln')
if ($SolutionRoot -and (Test-Path -Path $SolutionPath)) {
    $DllPath = [io.path]::Combine($SolutionRoot, "OfficeIMO.Word", "bin", "Debug", $DotNetVersion, "OfficeIMO.Word.dll")
    if (-not (Test-Path -Path $DllPath)) {
        Write-Host -Object "DLL not found at $DllPath" -ForegroundColor Red
        return
    }
    $DocsPath = [io.path]::Combine($SolutionRoot, "Docs")

    if (-not (Get-Command -Name xmldoc2md -ErrorAction SilentlyContinue)) {
        Write-Host -Object "xmldoc2md not found" -ForegroundColor Red
        return
    }
    xmldoc2md $DllPath $DocsPath
}