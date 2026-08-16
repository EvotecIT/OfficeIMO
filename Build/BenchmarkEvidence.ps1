function Get-BenchmarkEvidenceLocation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $ComparisonId,
        [Parameter(Mandatory)]
        [ValidateSet('windows', 'linux', 'macos')]
        [string] $Platform,
        [Parameter(Mandatory)]
        [ValidateSet('quick', 'full')]
        [string] $RunMode,
        [Parameter(Mandatory)]
        [string] $StaticRoot
    )

    $comparisonSlug = ($ComparisonId.Trim() -replace '[^A-Za-z0-9._-]', '-').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($comparisonSlug)) {
        throw 'ComparisonId must contain at least one filename-safe character.'
    }

    $fileName = "$comparisonSlug-$Platform-$RunMode.json"
    [pscustomobject]@{
        FileName = $fileName
        Path = Join-Path $StaticRoot $fileName
        ResultPath = "/data/benchmarks/library-comparisons/$fileName"
    }
}

function Get-BenchmarkEvidenceCaseIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object] $Row,
        [string[]] $VariableName
    )

    $scenario = [string] $Row.Scenario
    if ([string]::IsNullOrWhiteSpace($scenario)) {
        throw 'Benchmark evidence rows require a scenario identity.'
    }

    $ignoredVariables = @('namespace', 'type', 'fullname')
    $restrictVariables = $PSBoundParameters.ContainsKey('VariableName')
    $variables = @(
        if ($null -ne $Row.Variables) {
            foreach ($variable in $Row.Variables.GetEnumerator()) {
                $combined = '{0}={1}' -f ([string] $variable.Key), ([string] $variable.Value)
                foreach ($part in [regex]::Split($combined, '&(?=[^=&]+=)')) {
                    $equalsIndex = $part.IndexOf('=')
                    if ($equalsIndex -lt 1) {
                        continue
                    }

                    [pscustomobject]@{
                        Key = $part.Substring(0, $equalsIndex)
                        Value = $part.Substring($equalsIndex + 1)
                    }
                }
            }
        }
    )
    $variableParts = @(
        $variables |
            Where-Object {
                $key = ([string] $_.Key).ToLowerInvariant()
                $key -notin $ignoredVariables -and
                (-not $restrictVariables -or $key -in $VariableName)
            } |
            Sort-Object { [string] $_.Key } |
            ForEach-Object {
                '{0}={1}' -f ([string] $_.Key), ([string] $_.Value)
            }
    )

    if ($variableParts.Count -eq 0) {
        return $scenario
    }

    return '{0}|{1}' -f $scenario, ($variableParts -join '&')
}

function Write-BenchmarkEvidenceResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Path,
        [Parameter(Mandatory)]
        [object] $InputObject
    )

    $benchmarkJsonType = [Type]::GetType('PowerForge.BenchmarkJson, PowerForge', $true)
    $writeMethod = $benchmarkJsonType.GetMethods() |
        Where-Object { $_.Name -eq 'Write' -and $_.IsGenericMethodDefinition } |
        Select-Object -First 1
    if ($null -eq $writeMethod) {
        throw 'PowerForge.BenchmarkJson.Write<T> is unavailable.'
    }

    $value = $InputObject.PSObject.BaseObject
    $closedMethod = $writeMethod.MakeGenericMethod($value.GetType())
    $arguments = [object[]] @([string] $Path, $value)
    $null = $closedMethod.Invoke($null, $arguments)
}
