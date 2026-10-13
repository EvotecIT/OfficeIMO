<#
.SYNOPSIS
Runs the native engine's document journey through its C interface: open, read facts, render, rotate, undo, redo,
save a copy and reopen it. Use it on a published library (Windows .dll, macOS .dylib, Linux .so) to show the exports
work outside .NET.

.DESCRIPTION
The journey is the C program journey/officeimo_journey.c, compiled against include/officeimo_studio.h. It runs as a
native host because a .NET process cannot load the library on macOS or Linux: both runtimes handle the same
thread-suspension signal. The compiler is $env:CC when set, otherwise cc or clang, and on Windows clang or the
Visual C++ compiler found through vswhere.

.EXAMPLE
./Test-NativeEngine.ps1 -Library ./publish/OfficeIMOStudioEngine.dylib -Pdf ./sample.pdf -WorkDirectory /tmp/engine-check
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Library,
    [Parameter(Mandatory)][string]$Pdf,
    [Parameter(Mandatory)][string]$WorkDirectory
)
$ErrorActionPreference = 'Stop'
$Library = (Resolve-Path $Library).Path
New-Item -ItemType Directory -Force $WorkDirectory | Out-Null
$WorkDirectory = (Resolve-Path $WorkDirectory).Path
$source = Join-Path $WorkDirectory 'source.pdf'
Copy-Item $Pdf $source -Force

# A throwaway signing identity for the journey's certificate signature.
$key = [System.Security.Cryptography.RSA]::Create(2048)
$request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=OfficeIMO Journey', $key,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
$certificate = $request.CreateSelfSigned([DateTimeOffset]::Now.AddDays(-1), [DateTimeOffset]::Now.AddDays(30))
[IO.File]::WriteAllBytes((Join-Path $WorkDirectory 'signer.p12'),
    $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, 'journey'))
$certificate.Dispose()
$key.Dispose()

$journeySource = Join-Path $PSScriptRoot 'journey/officeimo_journey.c'
$include = Join-Path $PSScriptRoot 'include'
$journey = Join-Path $WorkDirectory $(if ($IsWindows) { 'officeimo_journey.exe' } else { 'officeimo_journey' })

function Find-Compiler {
    if ($env:CC) { return $env:CC }
    foreach ($name in $(if ($IsWindows) { 'clang', 'cl' } else { 'cc', 'clang' })) {
        $command = Get-Command $name -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { return $command.Source }
    }
    return $null
}

$compiler = Find-Compiler
if ($IsWindows -and (-not $compiler -or [IO.Path]::GetFileNameWithoutExtension($compiler) -eq 'cl')) {
    # The Visual C++ compiler needs its environment unless it is already on PATH. NativeAOT publishing on Windows
    # requires the same build tools.
    $setup = ''
    if (-not $compiler) {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = if (Test-Path $vswhere) {
            & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        }
        if (-not $installation) { throw 'No C compiler found. Install clang or the Visual C++ build tools, or set CC.' }
        $setup = "call `"$(Join-Path $installation 'VC\Auxiliary\Build\vcvars64.bat')`" >nul && "
    }
    cmd /c "$setup cl /nologo /std:clatest /O1 /I`"$include`" `"$journeySource`" /Fe:`"$journey`" /Fo:`"$WorkDirectory\\`""
} elseif ($compiler) {
    $arguments = @('-std=gnu11', '-Wall', '-O1', "-I$include", $journeySource, '-o', $journey)
    if ($IsLinux) { $arguments += '-ldl' }
    & $compiler @arguments
} else {
    throw 'No C compiler found. Install the Xcode command-line tools or a C compiler such as clang, or set CC.'
}
if ($LASTEXITCODE -ne 0) { throw "Compiling the journey failed with exit code $LASTEXITCODE." }

& $journey $Library $source $WorkDirectory
$journeyExit = $LASTEXITCODE
if ((Get-FileHash $Pdf).Hash -ne (Get-FileHash (Join-Path $WorkDirectory 'source.pdf')).Hash) {
    Write-Host 'FAILED: the source file changed; the journey saves only to a copy'
    exit 1
}
if ($journeyExit -ne 0) { exit $journeyExit }
