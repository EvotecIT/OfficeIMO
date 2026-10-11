<#
.SYNOPSIS
Reads command metadata in a clean PowerShell process, without build-tool assemblies.
#>
param(
    [Parameter(Mandatory)][string] $RepositoryRoot,
    [Parameter(Mandatory)][string] $ManifestPath,
    [Parameter(Mandatory)][string] $OutputPath
)

$ErrorActionPreference = 'Stop'
$moduleName = 'PSWriteOffice'
$manifest = Import-PowerShellDataFile -LiteralPath $ManifestPath
$expectedCommands = @($manifest.CmdletsToExport) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne '*' } | Sort-Object -Unique

function Get-RelativeFilePath {
    param(
        [Parameter(Mandatory)][string] $BasePath,
        [Parameter(Mandatory)][string] $Path
    )

    $baseFullPath = [System.IO.Path]::GetFullPath($BasePath)
    $separator = [System.IO.Path]::DirectorySeparatorChar.ToString()
    if (-not $baseFullPath.EndsWith($separator, [System.StringComparison]::Ordinal)) {
        $baseFullPath += $separator
    }

    $targetFullPath = [System.IO.Path]::GetFullPath($Path)
    $baseUri = [System.Uri] $baseFullPath
    $targetUri = [System.Uri] $targetFullPath
    if (-not [string]::Equals($baseUri.Scheme, $targetUri.Scheme, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $targetFullPath
    }

    $relativeUri = $baseUri.MakeRelativeUri($targetUri).ToString()
    return [System.Uri]::UnescapeDataString($relativeUri).Replace('/', [System.IO.Path]::DirectorySeparatorChar)
}

$sourceFilesByType = @{}
$sourceFilePriorityByType = @{}
foreach ($sourceFile in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'Sources\PSWriteOffice\Cmdlets') -Recurse -Filter '*.cs' -File) {
    $sourceText = Get-Content -LiteralPath $sourceFile.FullName -Raw -Encoding UTF8
    $containsCmdletDeclaration = $sourceText -match '(?m)^\s*\[(?:System\.Management\.Automation\.)?Cmdlet(?:Attribute)?\s*\('
    $typeNames = @($sourceFile.BaseName) + @(
        [regex]::Matches($sourceText, '\b(?:sealed\s+)?(?:partial\s+)?class\s+(?<name>[A-Za-z_][A-Za-z0-9_]*)\b') |
            ForEach-Object { $_.Groups['name'].Value }
    )

    foreach ($typeName in @($typeNames | Sort-Object -Unique)) {
        $priority = $(if ($containsCmdletDeclaration) { 2 } else { 0 }) + $(if ($sourceFile.BaseName -eq $typeName) { 1 } else { 0 })
        if (-not $sourceFilesByType.ContainsKey($typeName) -or $priority -gt $sourceFilePriorityByType[$typeName]) {
            $sourceFilesByType[$typeName] = $sourceFile.FullName
            $sourceFilePriorityByType[$typeName] = $priority
        }
    }
}

Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
try {
    Import-Module $ManifestPath -Force -ErrorAction Stop | Out-Null
    $allCommands = @(Get-Command -Module $moduleName -ErrorAction Stop)
    $commandsByName = @{}
    $aliasesByTarget = @{}

    foreach ($command in $allCommands) {
        $commandsByName[$command.Name] = $command
        if ($command.CommandType -ne 'Alias' -or [string]::IsNullOrWhiteSpace($command.ResolvedCommandName)) {
            continue
        }

        if (-not $aliasesByTarget.ContainsKey($command.ResolvedCommandName)) {
            $aliasesByTarget[$command.ResolvedCommandName] = [System.Collections.Generic.List[string]]::new()
        }
        [void] $aliasesByTarget[$command.ResolvedCommandName].Add($command.Name)
    }

    $missingCommands = @($expectedCommands | Where-Object { -not $commandsByName.ContainsKey($_) })
    if ($missingCommands.Count -gt 0) {
        throw "The built module does not export every manifest cmdlet: $($missingCommands -join ', ')"
    }

    $commandMetadata = foreach ($commandName in $expectedCommands) {
        $command = $commandsByName[$commandName]
        $declaredAliases = @()
        if ($command.ImplementingType) {
            $declaredAliases = @($command.ImplementingType.GetCustomAttributes([System.Management.Automation.AliasAttribute], $true) |
                ForEach-Object { $_.AliasNames })
        }
        $sourcePath = $null
        $sourceLine = $null
        $implementingType = $command.ImplementingType
        if ($implementingType -and $sourceFilesByType.ContainsKey($implementingType.Name)) {
            $sourceFullPath = $sourceFilesByType[$implementingType.Name]
            $sourcePath = (Get-RelativeFilePath -BasePath $RepositoryRoot -Path $sourceFullPath).Replace('\', '/')
            $typePattern = '\bclass\s+' + [regex]::Escape($implementingType.Name) + '\b'
            $typeMatch = Select-String -LiteralPath $sourceFullPath -Pattern $typePattern | Select-Object -First 1
            if ($typeMatch) {
                $sourceLine = $typeMatch.LineNumber
            }
        }

        $entry = [ordered]@{
            name = $commandName
            kind = 'Cmdlet'
            aliases = @(@($aliasesByTarget[$commandName]) + $declaredAliases |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique)
        }
        if ($sourcePath) {
            $entry.sourcePath = $sourcePath
            $entry.sourceLine = $sourceLine
            $entry.sourceUrl = "https://github.com/EvotecIT/PSWriteOffice/blob/main/$sourcePath#L$sourceLine"
        }
        $entry
    }
} finally {
    Remove-Module -Name $moduleName -Force -ErrorAction SilentlyContinue
}

@($commandMetadata) | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
