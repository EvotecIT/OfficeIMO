if (-not ('System.IO.Compression.ZipFile' -as [type])) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
}

function New-TestOfficeImageFile {
    param(
        [Parameter(Mandatory)]
        [string] $Directory,

        [string] $Name = 'OfficeIMO.bmp'
    )

    if (-not (Test-Path -Path $Directory)) {
        $null = New-Item -Path $Directory -ItemType Directory -Force
    }

    $path = Join-Path $Directory $Name
    [byte[]] $bytes = 0x42, 0x4D, 0x3A, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x36, 0x00, 0x00, 0x00, 0x28, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00,
        0x00, 0x00, 0x01, 0x00, 0x18, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x04, 0x00, 0x00, 0x00, 0x13, 0x0B,
        0x00, 0x00, 0x13, 0x0B, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF,
        0xFF, 0x00
    [System.IO.File]::WriteAllBytes($path, $bytes)

    $path
}

function New-TestTesseractExecutable {
    param(
        [Parameter(Mandatory)]
        [string] $Directory,

        [string] $CreateFilePath
    )

    $onWindows = $PSVersionTable.PSEdition -eq 'Desktop' -or $IsWindows
    if ($onWindows) {
        $path = Join-Path $Directory 'fake-tesseract.cmd'
        $content = @'
@echo off
if "%~1"=="--version" (
  echo tesseract 5.5.0
  exit /b 0
)
if "%~1"=="--list-langs" (
  echo List of available languages in fake runtime:
  echo eng
  echo pol
  exit /b 0
)
__SIDE_EFFECT__
set "output=%~2.tsv"
>"%output%" echo level	page_num	block_num	par_num	line_num	word_num	left	top	width	height	conf	text
>>"%output%" echo 5	1	1	1	1	1	10	12	120	24	96.0	OfficeIMO
>>"%output%" echo 5	1	1	1	1	2	140	12	72	24	94.0	OCR
exit /b 0
'@
        $sideEffect = if ($CreateFilePath) {
            '>' + '"' + $CreateFilePath + '" echo late destination'
        } else {
            ''
        }
        $content = $content.Replace('__SIDE_EFFECT__', $sideEffect) -replace "`r?`n", [Environment]::NewLine
        [IO.File]::WriteAllText($path, $content, [Text.Encoding]::ASCII)
        return $path
    }

    $path = Join-Path $Directory 'fake-tesseract'
    $content = @'
#!/usr/bin/env sh
if [ "$1" = "--version" ]; then
  echo "tesseract 5.5.0"
  exit 0
fi
if [ "$1" = "--list-langs" ]; then
  echo "List of available languages in fake runtime:"
  echo "eng"
  echo "pol"
  exit 0
fi
__SIDE_EFFECT__
printf 'level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n' > "${2}.tsv"
printf '5\t1\t1\t1\t1\t1\t10\t12\t120\t24\t96.0\tOfficeIMO\n' >> "${2}.tsv"
printf '5\t1\t1\t1\t1\t2\t140\t12\t72\t24\t94.0\tOCR\n' >> "${2}.tsv"
'@
    if ($CreateFilePath -and $CreateFilePath.Contains("'")) {
        throw 'The test side-effect path cannot contain a single quote.'
    }
    $sideEffect = if ($CreateFilePath) {
        "printf '%s' 'late destination' > '$CreateFilePath'"
    } else {
        ''
    }
    $content = $content.Replace('__SIDE_EFFECT__', $sideEffect) -replace "`r?`n", "`n"
    [IO.File]::WriteAllText($path, $content, [Text.Encoding]::ASCII)
    & chmod '+x' $path
    return $path
}

function Get-ZipXmlDocumentLocal {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string] $Entry
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $zipEntry = $archive.GetEntry($Entry)
        if (-not $zipEntry) {
            throw "Zip entry '$Entry' not found in '$Path'."
        }

        $stream = $zipEntry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream)
            try {
                return [xml] $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}

function Read-XlsxEntryText {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string] $Entry
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $zipEntry = $archive.GetEntry($Entry)
        if (-not $zipEntry) {
            throw "Zip entry '$Entry' not found in '$Path'."
        }

        $stream = $zipEntry.Open()
        try {
            $reader = [System.IO.StreamReader]::new($stream)
            try {
                return $reader.ReadToEnd()
            } finally {
                $reader.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}

function Set-XlsxEntryTextLocal {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [string] $Entry,

        [Parameter(Mandatory)]
        [string] $Text
    )

    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $zipEntry = $archive.GetEntry($Entry)
        if (-not $zipEntry) {
            throw "Zip entry '$Entry' not found in '$Path'."
        }

        $zipEntry.Delete()
        $newEntry = $archive.CreateEntry($Entry)
        $stream = $newEntry.Open()
        try {
            $encoding = [System.Text.UTF8Encoding]::new($false)
            $writer = [System.IO.StreamWriter]::new($stream, $encoding)
            try {
                $writer.Write($Text)
            } finally {
                $writer.Dispose()
            }
        } finally {
            $stream.Dispose()
        }
    } finally {
        $archive.Dispose()
    }
}

function Get-ZipEntriesLocal {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            $entry.FullName
        }
    } finally {
        $archive.Dispose()
    }
}

function Get-TestPowerPointSlideEntry {
    param(
        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter(Mandatory)]
        [int] $Index
    )

    $entries = @(Get-ZipEntriesLocal -Path $Path |
        Where-Object { $_ -match '^ppt/slides/slide(\d+)\.xml$' } |
        Sort-Object { [int] ([regex]::Match($_, 'slide(\d+)\.xml$').Groups[1].Value) })

    if ($Index -lt 0 -or $Index -ge $entries.Count) {
        throw "PowerPoint slide index $Index was not found in '$Path'."
    }

    $entries[$Index]
}

function Start-TestHttpFileServer {
    param(
        [Parameter(Mandatory)]
        [string] $FilePath,

        [string] $ContentType = 'application/octet-stream',

        [ValidateRange(1, 10000)]
        [int] $RequestCount = 1
    )

    if (-not ('OfficeIMOTestHttpFileServer' -as [type])) {
        Add-Type -Path (Join-Path $PSScriptRoot 'Support/HttpFileServer.cs')
    }

    $server = [OfficeIMOTestHttpFileServer]::new($FilePath, $ContentType, $RequestCount)
    $fileName = [uri]::EscapeDataString([IO.Path]::GetFileName($FilePath))
    [pscustomobject]@{
        Url = "http://127.0.0.1:$($server.Port)/$fileName"
        Server = $server
    }
}

function Stop-TestHttpFileServer {
    param(
        [Parameter(Mandatory)]
        $Server
    )

    $Server.Server.Dispose()
    if ($Server.Server.Failure) {
        throw $Server.Server.Failure
    }
}

function Get-TestPSWriteOfficeAssemblyLoadContext {
    param(
        [string] $CommandName = 'New-OfficeWord'
    )

    $assemblyLoadContextType = 'System.Runtime.Loader.AssemblyLoadContext' -as [type]
    if (-not $assemblyLoadContextType) {
        return $null
    }

    $command = Get-Command -Name $CommandName -ErrorAction Stop
    $assemblyLoadContextType.GetMethod('GetLoadContext').Invoke($null, @($command.ImplementingType.Assembly))
}

function Get-TestPSWriteOfficeAssembly {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [string] $CommandName = 'New-OfficeWord'
    )

    $loadContext = Get-TestPSWriteOfficeAssemblyLoadContext -CommandName $CommandName
    if ($loadContext) {
        return $loadContext.LoadFromAssemblyName([System.Reflection.AssemblyName]::new($Name))
    }

    $assembly = [AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq $Name } |
        Select-Object -First 1
    if ($assembly) {
        return $assembly
    }

    [System.Reflection.Assembly]::Load($Name)
}

function Get-TestPSWriteOfficeType {
    param(
        [Parameter(Mandatory)]
        [string] $AssemblyName,

        [Parameter(Mandatory)]
        [string] $TypeName,

        [string] $CommandName = 'New-OfficeWord'
    )

    $assembly = Get-TestPSWriteOfficeAssembly -Name $AssemblyName -CommandName $CommandName
    $assembly.GetType($TypeName, $true)
}

function Get-TestPSWriteOfficeEnumValue {
    param(
        [Parameter(Mandatory)]
        [string] $AssemblyName,

        [Parameter(Mandatory)]
        [string] $TypeName,

        [Parameter(Mandatory)]
        [string] $Name,

        [string] $CommandName = 'New-OfficeWord'
    )

    $type = Get-TestPSWriteOfficeType -AssemblyName $AssemblyName -TypeName $TypeName -CommandName $CommandName
    [Enum]::Parse($type, $Name)
}
