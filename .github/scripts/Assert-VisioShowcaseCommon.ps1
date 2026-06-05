function Get-RequiredArray {
    param(
        [Parameter(Mandatory = $true)]
        [object] $Value
    )

    if ($null -eq $Value) {
        return @()
    }

    return @($Value)
}

function Resolve-ShowcaseArtifactPath {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath
    )

    Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($RelativePath)) -Message 'Artifact relativePath cannot be empty.'
    Assert-Condition -Condition (-not [System.IO.Path]::IsPathRooted($RelativePath)) -Message "Artifact path must be relative: $RelativePath"

    $combined = $script:ShowcaseRoot
    foreach ($part in ($RelativePath -split '/')) {
        Assert-Condition -Condition (-not [string]::IsNullOrWhiteSpace($part)) -Message "Artifact path contains an empty segment: $RelativePath"
        Assert-Condition -Condition ($part -ne '..') -Message "Artifact path cannot escape the showcase root: $RelativePath"
        $combined = Join-Path $combined $part
    }

    $fullPath = [System.IO.Path]::GetFullPath($combined)
    Assert-Condition -Condition ($fullPath.StartsWith($script:ShowcaseRootWithSeparator, [StringComparison]::OrdinalIgnoreCase)) -Message "Artifact path escapes the showcase root: $RelativePath"
    return $fullPath
}

function ConvertTo-GalleryHref {
    param(
        [Parameter(Mandatory = $true)]
        [string] $RelativePath
    )

    return (($RelativePath -split '/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
}

function ConvertTo-GalleryFragmentId {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Value
    )

    $fragment = ($Value.ToLowerInvariant() -replace '[^a-z0-9]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($fragment)) {
        return 'item'
    }

    return $fragment
}
