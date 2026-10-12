<#
.SYNOPSIS
Runs the native engine's document journey through its C interface: open, read facts, render, rotate, undo, redo
by rotating again, save a copy and reopen it. Use it on a published library (Windows .dll, macOS .dylib, Linux .so)
to show the exports work outside .NET.

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

Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class OfficeImoEngine {
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_abi_version")] public static extern int AbiVersion();
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_free")] public static extern void Free(IntPtr memory);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_initialize")] public static extern int Initialize(byte[] directory, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_cancellation_create")] public static extern IntPtr CancellationCreate();
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_cancellation_cancel")] public static extern void CancellationCancel(IntPtr cancellation);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_cancellation_free")] public static extern void CancellationFree(IntPtr cancellation);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_open")] public static extern int Open(byte[] path, byte[] password, IntPtr cancellation, out IntPtr document, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_close")] public static extern void Close(IntPtr document);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_info_json")] public static extern int Info(IntPtr document, out IntPtr json, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_render_page")] public static extern int Render(IntPtr document, int page, double scale, IntPtr cancellation, out IntPtr png, out UIntPtr length, out int width, out int height, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_rotate_pages")] public static extern int Rotate(IntPtr document, int[] pages, int count, int degrees, IntPtr cancellation, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_undo")] public static extern int Undo(IntPtr document, IntPtr cancellation, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_redo")] public static extern int Redo(IntPtr document, IntPtr cancellation, out IntPtr error);
    [DllImport("$($Library.Replace('\','\\'))", EntryPoint = "officeimo_document_save")] public static extern int Save(IntPtr document, byte[] path, IntPtr cancellation, out IntPtr error);
    public static byte[] Utf8(string value) { return value == null ? null : Encoding.UTF8.GetBytes(value + "\0"); }
    public static string Take(IntPtr text) { if (text == IntPtr.Zero) return null; string value = Marshal.PtrToStringUTF8(text); Free(text); return value; }
}
"@

$failures = [Collections.Generic.List[string]]::new()
function Check([string]$Step, [int]$Status, [IntPtr]$Message, [int]$Expected = 0) {
    $text = [OfficeImoEngine]::Take($Message)
    $line = "{0,-28} status {1}{2}" -f $Step, $Status, $(if ($text) { " - $text" } else { '' })
    Write-Host $line
    if ($Status -ne $Expected) { $failures.Add($line) }
}
function Info([IntPtr]$Document) {
    $json = [IntPtr]::Zero; $err = [IntPtr]::Zero
    Check 'info' ([OfficeImoEngine]::Info($Document, [ref]$json, [ref]$err)) $err
    [OfficeImoEngine]::Take($json) | ConvertFrom-Json
}

$err = [IntPtr]::Zero; $document = [IntPtr]::Zero
Write-Host "ABI version $([OfficeImoEngine]::AbiVersion())"
Check 'open before initialize' ([OfficeImoEngine]::Open([OfficeImoEngine]::Utf8($source), $null, 0, [ref]$document, [ref]$err)) $err 8
Check 'initialize' ([OfficeImoEngine]::Initialize([OfficeImoEngine]::Utf8((Join-Path $WorkDirectory 'data')), [ref]$err)) $err
$clock = [Diagnostics.Stopwatch]::StartNew()
Check 'open' ([OfficeImoEngine]::Open([OfficeImoEngine]::Utf8($source), $null, 0, [ref]$document, [ref]$err)) $err
Write-Host "  opened in $($clock.ElapsedMilliseconds) ms"
$info = Info $document
Write-Host "  $($info.fileName): $($info.pages.Count) page(s), first $($info.pages[0].width) x $($info.pages[0].height) pt"

$png = [IntPtr]::Zero; $length = [UIntPtr]::Zero; $width = 0; $height = 0
$clock.Restart()
Check 'render page 1' ([OfficeImoEngine]::Render($document, 1, 1.5, 0, [ref]$png, [ref]$length, [ref]$width, [ref]$height, [ref]$err)) $err
Write-Host "  $width x $height px, $length bytes in $($clock.ElapsedMilliseconds) ms"
$bytes = New-Object byte[] ([int]$length.ToUInt64())
[Runtime.InteropServices.Marshal]::Copy($png, $bytes, 0, $bytes.Length); [OfficeImoEngine]::Free($png)
[IO.File]::WriteAllBytes((Join-Path $WorkDirectory 'page-1.png'), $bytes)
Check 'render missing page' ([OfficeImoEngine]::Render($document, 9999, 1.0, 0, [ref]$png, [ref]$length, [ref]$width, [ref]$height, [ref]$err)) $err 3

$cancellation = [OfficeImoEngine]::CancellationCreate(); [OfficeImoEngine]::CancellationCancel($cancellation)
Check 'cancelled render' ([OfficeImoEngine]::Render($document, 1, 1.0, $cancellation, [ref]$png, [ref]$length, [ref]$width, [ref]$height, [ref]$err)) $err 2
[OfficeImoEngine]::CancellationFree($cancellation)

Check 'rotate page 1' ([OfficeImoEngine]::Rotate($document, [int[]]@(1), 1, 90, 0, [ref]$err)) $err
$info = Info $document
if ($info.pages[0].rotation -ne 90 -or -not $info.isDirty -or -not $info.canUndo) { $failures.Add("rotation not recorded: $($info | ConvertTo-Json -Compress)") }
Check 'undo' ([OfficeImoEngine]::Undo($document, 0, [ref]$err)) $err
$info = Info $document
if ($info.pages[0].rotation -ne 0 -or -not $info.canRedo) { $failures.Add("undo not applied: $($info | ConvertTo-Json -Compress)") }
Check 'redo' ([OfficeImoEngine]::Redo($document, 0, [ref]$err)) $err
$copy = Join-Path $WorkDirectory 'rotated.pdf'
Check 'save copy' ([OfficeImoEngine]::Save($document, [OfficeImoEngine]::Utf8($copy), 0, [ref]$err)) $err
[OfficeImoEngine]::Close($document)

$reopened = [IntPtr]::Zero
Check 'reopen copy' ([OfficeImoEngine]::Open([OfficeImoEngine]::Utf8($copy), $null, 0, [ref]$reopened, [ref]$err)) $err
$info = Info $reopened
if ($info.pages[0].rotation -ne 90) { $failures.Add("saved copy has rotation $($info.pages[0].rotation)") }
[OfficeImoEngine]::Close($reopened)
if ((Get-FileHash $source).Hash -ne (Get-FileHash $Pdf).Hash) { $failures.Add('the source file changed') }

if ($failures.Count -gt 0) { $failures | ForEach-Object { Write-Host "FAILED: $_" }; exit 1 }
Write-Host 'Native engine journey passed.'
