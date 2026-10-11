BeforeDiscovery {
    $repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
    $examplesRoot = Join-Path $repoRoot 'Examples'
    $exampleCases = Get-ChildItem -LiteralPath $examplesRoot -Recurse -File -Filter '*.ps1' |
        Sort-Object FullName |
        ForEach-Object {
            $relativePath = $_.FullName.Substring($repoRoot.Length)
            $relativePath = $relativePath.TrimStart([char[]] @(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar
            ))

            @{
                Path = $_.FullName
                RelativePath = $relativePath
            }
        }

}

Describe 'Repository example scripts' {
    It 'parses <RelativePath>' -ForEach $exampleCases {
        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            $Path,
            [ref] $tokens,
            [ref] $parseErrors
        ) | Out-Null

        @($parseErrors | ForEach-Object Message) | Should -BeNullOrEmpty
    }

}
