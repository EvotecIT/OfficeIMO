param([Parameter(Mandatory)][string] $RepositoryRoot)

# One project classification is shared by catalog generation and source validation.
$allProjects = @(Get-ChildItem -LiteralPath $RepositoryRoot -Recurse -File -Filter '*.csproj' |
    Where-Object {
        $relativeProjectPath = [System.IO.Path]::GetRelativePath($RepositoryRoot, $_.FullName).Replace('\', '/')
        $relativeProjectPath -notmatch '(^|/)(?:bin|obj)(?:/|$)' -and
            $relativeProjectPath -notmatch '(^|/)(?:\.ci-artifacts|\.playwright-cli|\.powerforge-runner|\.?artifacts|Ignore|_worktrees)(?:/|$)' -and
            $relativeProjectPath -notmatch '^Website/projects/'
    })

$testProjects = @($allProjects | Where-Object { $_.BaseName -match '(?:^|\.)(?:Tests|VerifyTests)(?:\.|$)' })
$benchmarkProjects = @($allProjects | Where-Object { $_.BaseName -match '(?:^|\.)Benchmarks(?:\.|$)' })
$validationProjects = @($allProjects | Where-Object {
    $relativeProjectPath = [System.IO.Path]::GetRelativePath($RepositoryRoot, $_.FullName).Replace('\', '/')
    $relativeProjectPath -match '(^|/)Build/' -or
    $_.BaseName -match '(?:^|\.)AotSmoke$' -or
    $_.BaseName -in @(
        'OfficeIMO.Examples',
        'OfficeIMO.MarkdownRenderer.SamplePlugin',
        'OfficeIMO.Project.Verification',
        'OfficeIMO.Project.IndependentVerification',
        'OfficeIMO.Project.ReportVerification'
    ) -or
    $relativeProjectPath -match '^Website/Apps/'
})
$productionProjects = @($allProjects | Where-Object {
    $_ -notin $testProjects -and
    $_ -notin $benchmarkProjects -and
    $_ -notin $validationProjects
})

[PSCustomObject]@{
    All = $allProjects
    Tests = $testProjects
    Benchmarks = $benchmarkProjects
    Validation = $validationProjects
    Production = $productionProjects
}
