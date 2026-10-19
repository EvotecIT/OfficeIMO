Set-StrictMode -Version Latest

$script:ExcelBenchmarkModuleRoot = $null
$script:ExcelBenchmarkPackageRoot = $null
$script:ExcelBenchmarkInitializedEngines = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Get-ExcelBenchmarkDefaultRowCount {
    param(
        [Parameter(Mandatory)]
        [string] $Suite
    )

    switch ($Suite) {
        Smoke { @(1000) }
        Standard { @(1000, 10000, 25000) }
        Large { @(25000, 100000, 250000) }
        Full { @(1000, 10000, 25000, 100000) }
        SuperLarge { @(250000, 500000, 1000000) }
        default { @(1000, 10000, 25000) }
    }
}

function Assert-ExcelBenchmarkRowCount {
    param([int[]] $RowCount)
    @(
        foreach ($count in $RowCount) {
            if ($count -le 0) { throw "RowCount must be greater than zero. Value: $count" }
            $count
        }
    )
}

function Get-ExcelBenchmarkIterationCount {
    param([string] $Suite)

    switch ($Suite) {
        Smoke { 5 }
        Standard { 5 }
        Large { 3 }
        Full { 5 }
        SuperLarge { 3 }
        default { 5 }
    }
}

function Get-CsvBenchmarkIterationCount {
    param([string] $Suite)

    switch ($Suite) {
        Smoke { 25 }
        Standard { 25 }
        Large { 11 }
        Full { 25 }
        SuperLarge { 7 }
        default { 25 }
    }
}

function Get-ExcelBenchmarkWarmupCount {
    param([string] $Suite)

    switch ($Suite) {
        Smoke { 3 }
        Standard { 3 }
        Large { 2 }
        Full { 2 }
        SuperLarge { 1 }
        default { 3 }
    }
}

function New-ExcelBenchmarkCase {
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter(Mandatory)]
        [string] $Label,

        [Parameter(Mandatory)]
        [string[]] $Suites,

        [Parameter(Mandatory)]
        [string] $OperationKey,

        [Parameter(Mandatory)]
        [string] $Profile,

        [string] $FileExtension = '.xlsx',

        [bool] $ValidateWorkbook = $true
    )

    $supportedEngines = @('PSWriteOffice')
    if ($OperationKey -notin @(
            'WriteCsv',
            'ReadCsvSource',
            'ReadUsedRangeDataTable',
            'ReadFullSheetDataTable',
            'ReadRangeDataTable',
            'ReadFullSheetDataReader',
            'ReadRangeDataReader',
            'ReadTableMetadata',
            'ReadNamedRangeMetadata'
        ) -and
        $Name -notin @('dataset-worksheets', 'datareader-table')) {
        $supportedEngines += 'ImportExcel'
    }
    if ($Name -in @('text-objects-default', 'import-default-full', 'import-default-range', 'read-no-header-range')) {
        $supportedEngines += 'ExcelFast'
    }

    [pscustomobject]@{
        Name = $Name
        Label = $Label
        Suites = $Suites -join ','
        OperationKey = $OperationKey
        DataProfile = $Profile
        FileExtension = $FileExtension
        ValidateWorkbook = $ValidateWorkbook
        SupportedEngines = $supportedEngines -join ','
    }
}

function Get-ExcelBenchmarkCase {
    param([string] $Suite)

    $basic = @('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')
    $table = @('Smoke', 'Standard', 'Large', 'Full')
    $standard = @('Standard', 'Large', 'Full')
    $scale = @('Standard', 'Large', 'Full', 'SuperLarge')
    $workflow = @('Standard', 'Large', 'Full')
    $report = @('Smoke', 'Standard', 'Large', 'Full')
    $dataSet = @('Large', 'Full')

    @(
        New-ExcelBenchmarkCase -Name csv-to-excel -Label 'Create workbook from CSV source' -Suites $workflow -OperationKey CsvToExcel -Profile MixedObjects
        New-ExcelBenchmarkCase -Name objects-table -Label 'Export objects as table' -Suites $table -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name objects-default -Label 'Export objects default' -Suites $basic -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name text-objects-default -Label 'Export text-only objects default' -Suites $basic -OperationKey WriteWorkbook -Profile TextObjects
        New-ExcelBenchmarkCase -Name objects-no-table -Label 'Export objects without a table' -Suites $scale -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name objects-table-autofit -Label 'Export objects as autofit table' -Suites $standard -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name objects-title-freeze -Label 'Export objects with title and frozen header' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name wide-objects-default -Label 'Export wide objects default' -Suites $scale -OperationKey WriteWorkbook -Profile WideObjects
        New-ExcelBenchmarkCase -Name datatable-default -Label 'Export DataTable default' -Suites $scale -OperationKey WriteWorkbook -Profile DataTable
        New-ExcelBenchmarkCase -Name datareader-table -Label 'Stream IDataReader as table' -Suites $scale -OperationKey WriteWorkbook -Profile DataTable
        New-ExcelBenchmarkCase -Name import-default-full -Label 'Read full sheet from default export' -Suites $basic -OperationKey ReadFullSheet -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name import-default-range -Label 'Read A1 range from default export' -Suites $scale -OperationKey ReadRange -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name import-datatable-full -Label 'Read full sheet as DataTable' -Suites $scale -OperationKey ReadFullSheetDataTable -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name import-datatable-range -Label 'Read A1 range as DataTable' -Suites $scale -OperationKey ReadRangeDataTable -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name import-datareader-full -Label 'Consume full sheet as IDataReader' -Suites $scale -OperationKey ReadFullSheetDataReader -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name import-datareader-range -Label 'Consume A1 range as IDataReader' -Suites $scale -OperationKey ReadRangeDataReader -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name read-no-header-range -Label 'Read selected range without headers' -Suites $standard -OperationKey ReadNoHeaderRange -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name read-used-range-datatable -Label 'Read used range as DataTable' -Suites $standard -OperationKey ReadUsedRangeDataTable -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name read-table-metadata -Label 'Read workbook table metadata' -Suites $standard -OperationKey ReadTableMetadata -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name multi-sheet-regions -Label 'Export regional workbook with one table per sheet' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name summary-formulas -Label 'Export data workbook with summary formulas' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name append-existing-table -Label 'Append rows to an existing workbook table' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name update-existing-workbook -Label 'Update cells and formulas in an existing workbook' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name many-small-sheets -Label 'Export many small worksheets' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name workbook-package-merge -Label 'Merge workbook sheets with package copy' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name named-range-workbook -Label 'Export workbook with named data range' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name read-named-range-metadata -Label 'Read workbook named range metadata' -Suites $standard -OperationKey ReadNamedRangeMetadata -Profile MixedObjects -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name chart-only-workbook -Label 'Export workbook with table and chart' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name pivot-only-workbook -Label 'Export workbook with table and pivot' -Suites $workflow -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name report-workbook -Label 'Export report workbook with table chart pivot formatting' -Suites $report -OperationKey WriteWorkbook -Profile MixedObjects
        New-ExcelBenchmarkCase -Name dataset-worksheets -Label 'Export DataSet worksheets' -Suites $dataSet -OperationKey WriteWorkbook -Profile DataSet
    ) | Where-Object { (($_.Suites -split ',') -contains $Suite) }
}

function Get-CsvBenchmarkCase {
    param([string] $Suite)

    $csv = @('Smoke', 'Standard', 'Large', 'Full', 'SuperLarge')
    @(
        foreach ($csvProfile in @(
            [pscustomobject]@{ Profile = 'MixedObjects'; Label = 'mixed' }
            [pscustomobject]@{ Profile = 'CsvQuotedObjects'; Label = 'quoted' }
            [pscustomobject]@{ Profile = 'CsvMultilineObjects'; Label = 'multiline' }
            [pscustomobject]@{ Profile = 'CsvWideObjects'; Label = 'wide' }
        )) {
            New-ExcelBenchmarkCase -Name ('csv-write-{0}' -f $csvProfile.Label) -Label ('Write CSV file ({0})' -f $csvProfile.Label) -Suites $csv -OperationKey WriteCsv -Profile $csvProfile.Profile -FileExtension '.csv' -ValidateWorkbook:$false
            New-ExcelBenchmarkCase -Name ('csv-write-gzip-{0}' -f $csvProfile.Label) -Label ('Write GZip CSV file ({0})' -f $csvProfile.Label) -Suites $csv -OperationKey WriteCsvGZip -Profile $csvProfile.Profile -FileExtension '.csv.gz' -ValidateWorkbook:$false
            New-ExcelBenchmarkCase -Name ('csv-read-source-{0}' -f $csvProfile.Label) -Label ('Read CSV file ({0})' -f $csvProfile.Label) -Suites $csv -OperationKey ReadCsvSource -Profile $csvProfile.Profile -FileExtension '.csv' -ValidateWorkbook:$false
            New-ExcelBenchmarkCase -Name ('csv-read-datatable-{0}' -f $csvProfile.Label) -Label ('Read CSV file as DataTable ({0})' -f $csvProfile.Label) -Suites $csv -OperationKey ReadCsvDataTable -Profile $csvProfile.Profile -FileExtension '.csv' -ValidateWorkbook:$false
            New-ExcelBenchmarkCase -Name ('csv-read-gzip-datatable-{0}' -f $csvProfile.Label) -Label ('Read GZip CSV file as DataTable ({0})' -f $csvProfile.Label) -Suites $csv -OperationKey ReadCsvGZipDataTable -Profile $csvProfile.Profile -FileExtension '.csv.gz' -ValidateWorkbook:$false
        }
        New-ExcelBenchmarkCase -Name csv-write-datatable -Label 'Write DataTable as CSV file' -Suites $csv -OperationKey WriteCsvDataTable -Profile DataTable -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-quick-single-column -Label 'dbatools QuickTest read first column' -Suites $csv -OperationKey ReadCsvQuickSingleColumn -Profile DbatoolsQuickCsv -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-quick-all-columns -Label 'dbatools QuickTest read all columns' -Suites $csv -OperationKey ReadCsvQuickAllColumns -Profile DbatoolsQuickCsv -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-wide-single-column -Label 'dbatools wide read first column' -Suites $csv -OperationKey ReadCsvQuickSingleColumn -Profile DbatoolsWideCsv -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-wide-all-columns -Label 'dbatools wide read all columns' -Suites $csv -OperationKey ReadCsvQuickAllColumns -Profile DbatoolsWideCsv -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-quoted-single-column -Label 'dbatools quoted read first column' -Suites $csv -OperationKey ReadCsvQuickSingleColumn -Profile DbatoolsQuotedCsv -FileExtension '.csv' -ValidateWorkbook:$false
        New-ExcelBenchmarkCase -Name csv-dbatools-quoted-all-columns -Label 'dbatools quoted read all columns' -Suites $csv -OperationKey ReadCsvQuickAllColumns -Profile DbatoolsQuotedCsv -FileExtension '.csv' -ValidateWorkbook:$false
    ) | Where-Object { (($_.Suites -split ',') -contains $Suite) }
}

function Test-ExcelBenchmarkEngineCaseSupport {
    param(
        [Parameter(Mandatory)]
        [string] $Engine,

        [Parameter(Mandatory)]
        [object] $Case
    )

    return ([string]$Case.SupportedEngines -split ',') -contains $Engine
}

function Test-CsvBenchmarkEngineSupport {
    param(
        [Parameter(Mandatory)]
        [string] $Engine,

        [Parameter(Mandatory)]
        [object] $Case
    )

    switch ($Engine) {
        PSWriteOffice { return $true }
        NativeCsv { return [string]$Case.OperationKey -in @('WriteCsv', 'WriteCsvGZip', 'WriteCsvDataTable', 'ReadCsvSource', 'ReadCsvDataTable', 'ReadCsvGZipDataTable', 'ReadCsvQuickSingleColumn', 'ReadCsvQuickAllColumns') }
        default { return $false }
    }
}

function Initialize-ExcelBenchmarkEngine {
    param(
        [Parameter(Mandatory)]
        [string] $Engine,

        [Parameter(Mandatory)]
        [object] $Run
    )

    if ($null -eq $script:ExcelBenchmarkInitializedEngines) {
        $script:ExcelBenchmarkInitializedEngines = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    }

    if (-not $script:ExcelBenchmarkModuleRoot) {
        $script:ExcelBenchmarkModuleRoot = Join-Path $Run.OutputRoot 'Modules'
        $script:ExcelBenchmarkPackageRoot = Join-Path $Run.OutputRoot 'Packages'
        New-Item -ItemType Directory -Force -Path $script:ExcelBenchmarkModuleRoot, $script:ExcelBenchmarkPackageRoot | Out-Null
        if (-not (($env:PSModulePath -split [IO.Path]::PathSeparator) -contains $script:ExcelBenchmarkModuleRoot)) {
            $env:PSModulePath = $script:ExcelBenchmarkModuleRoot + [IO.Path]::PathSeparator + $env:PSModulePath
        }
    }

    if ($script:ExcelBenchmarkInitializedEngines.Contains($Engine)) {
        return
    }

    switch ($Engine) {
        ImportExcel {
            if (-not (Get-Module -ListAvailable ImportExcel | Sort-Object Version -Descending | Select-Object -First 1)) {
                if ([bool]$Run.SkipImportExcelInstall) {
                    throw 'ImportExcel is not installed. Rerun without -SkipImportExcelInstall to save it under the benchmark module folder.'
                }
                Save-Module -Name ImportExcel -Path $script:ExcelBenchmarkModuleRoot -Repository PSGallery -Force
            }
            Import-Module ImportExcel -Force -ErrorAction Stop
        }
        ExcelFast {
            if (-not (Get-Module -ListAvailable ExcelFast | Sort-Object Version -Descending | Select-Object -First 1)) {
                throw 'ExcelFast is not available. Use Compare-ExcelPerformance.ps1 to prepare optional ExcelFast lanes, or run without the ExcelFast engine.'
            }
            Import-Module ExcelFast -Force -ErrorAction Stop
        }
    }

    [void]$script:ExcelBenchmarkInitializedEngines.Add($Engine)
}
