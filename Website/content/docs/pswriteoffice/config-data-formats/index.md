---
title: "Configuration Files and Data Conversion"
description: "Edit YAML without losing comments, read and write TOML, and convert between JSON, JSON Lines, YAML, TOML, XML, CSV, TSV, and Excel from PowerShell, with a report of anything the target cannot hold."
layout: docs
---

Configuration and data files are a large part of everyday scripting: Helm values, CI settings, Cargo and Poetry lock files, inventory exports, API responses. PSWriteOffice handles them with OfficeIMO's own YAML, TOML, CSV, and data-conversion engines, so none of this needs a second module or an external tool.

| Task | Commands |
| --- | --- |
| Change a YAML file and keep its comments and layout | `Open-OfficeYaml`, `Set-OfficeYamlNode`, `Remove-OfficeYamlNode`, `Save-OfficeYaml` |
| Layer one YAML file over another | `Merge-OfficeDocument` |
| Read or write YAML and TOML as objects | `Import-OfficeYaml`, `Export-OfficeYaml`, `Import-OfficeToml`, `Export-OfficeToml` and the `ConvertFrom-`/`ConvertTo-` forms |
| Check YAML before a deployment | `Test-OfficeYaml` |
| Convert a file between data formats | `Convert-OfficeDocument` |
| Write CSV with explicit quoting rules | `Export-OfficeCsv`, `ConvertTo-OfficeCsv`, `Import-OfficeCsv` |

## Edit YAML without rewriting it

`ConvertFrom-Yaml` and `ConvertTo-Yaml` style round trips lose comments, quoting, and key order. `Open-OfficeYaml` loads the file as a document instead, and saving it changes only the text you edited.

```powershell
Open-OfficeYaml -Path .\values.yaml {
    Set-OfficeYamlNode -Select 'image.tag' -Value '1.26'
    Set-OfficeYamlNode -Select 'resources.limits' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' })
}
```

Given a hand-maintained file, the comments, the single-quoted tag, and the flow-style list stay exactly as they were. The new `resources` block follows the file's indentation:

```yaml
# Defaults for the web chart
image:
  repository: nginx   # pinned by CI
  tag: '1.26'
replicas: 1

# Ports the service exposes
ports: [80, 443]
resources:
  limits:
    cpu: 500m
    memory: 256Mi
```

The script block saves and closes the document when it ends. Use `Open-OfficeYaml` without a block, `Save-OfficeYaml`, and `Close-OfficeYaml` when the document needs to stay open across several steps. `Get-OfficeYamlNode` reads a value by the same path grammar.

## Merge overlays

`Merge-OfficeDocument` layers YAML files in order, the way Helm combines values files. Later files win; the base file keeps its comments and layout. `-PassThruReport` returns what changed:

```powershell
Merge-OfficeDocument -Path .\values.yaml, .\values.prod.yaml -OutputPath .\values.merged.yaml -PassThruReport |
    Select-Object -Property Added, Replaced, Removed
```

```text
Added Replaced              Removed
----- --------              -------
{}    {image.tag, replicas} {}
```

## Validate before deploying

`Test-OfficeYaml` returns `$true` or `$false`. Add `-Strict` to reject duplicate keys, which YAML otherwise resolves as last-wins with a warning:

```powershell
if (-not (Test-OfficeYaml -Path .\values.merged.yaml -Strict)) {
    throw 'values.merged.yaml is not valid YAML.'
}
```

Add `-Detailed` when a person needs to fix the file. The result carries a stable error code, the line and column, and the offending text:

```powershell
Test-OfficeYaml -Path .\broken.yaml -Detailed | Select-Object IsValid, Code, Line, Column
```

```text
IsValid Code                         Line Column
------- ----                         ---- ------
  False yaml-syntax-unexpected-token    3      2
```

Use `-Untrusted` on `Import-OfficeYaml` and `Import-OfficeToml` for files from outside your pipeline. It applies tighter limits on depth, size, and alias expansion.

## Read and write TOML

The TOML commands follow the `ConvertFrom-Json` and `ConvertTo-Json` conventions. Tables become objects and arrays of tables become arrays:

```powershell
$lock = Import-OfficeToml -Path .\Cargo.lock
$lock.package | Select-Object -Property name, version
```

Writing works from ordinary hashtables and objects. Nested objects become `[table]` sections, and date-time values keep their TOML type:

```powershell
[ordered]@{
    title   = 'Nightly report'
    retries = 3
    owner   = [ordered]@{ team = 'Ops'; updated = [DateTimeOffset]'2026-11-11T09:30:00Z' }
} | Export-OfficeToml -Path .\settings.toml
```

```toml
title = "Nightly report"
retries = 3

[owner]
team = "Ops"
updated = 2026-11-11T09:30:00Z
```

TOML has no null. A `$null` property is left out and reported with a warning rather than written as an empty string.

## Convert between data formats

`Convert-OfficeDocument` converts between JSON, JSON Lines (`.jsonl`, `.ndjson`), YAML, TOML, XML, CSV, TSV, and Excel worksheets. The output extension selects the target, and the data streams through without being loaded into an object model first:

```powershell
Convert-OfficeDocument -Path .\checks.jsonl -OutputPath .\checks.csv
Convert-OfficeDocument -Path .\checks.jsonl -OutputPath .\checks.xlsx
Convert-OfficeDocument -Path .\Cargo.lock -OutputPath .\packages.csv -Options @{ Select = 'package' }
```

`Select` picks the rows from a nested document. Other `-Options` keys include `RootName` for XML, `Columns` to fix the CSV column order, and `SampleRows` to control how many rows decide the CSV columns.

### See what the target could not hold

Every format has limits: TOML has no null, CSV has no nesting, XML attributes have no JSON equivalent, and Excel cells hold 15 significant digits. The result object reports each case with a stable code and the location:

```powershell
$result = Convert-OfficeDocument -Path .\service.json -OutputPath .\service.toml
$result.HasLoss
$result.Diagnostics | Format-List Code, Message, Location, LossKind
```

```text
True

Code     : null-dropped
Message  : TOML has no null, so null values were omitted.
Location : service.owner
LossKind : Omission
```

Add `-FailOnLoss` in automation where a silent change is worse than a failed job. The conversion throws `OfficeConversionException`, and no output file is written:

```powershell
Convert-OfficeDocument -Path .\service.json -OutputPath .\service.toml -FailOnLoss
```

## CSV with explicit rules

`Export-OfficeCsv` and `ConvertTo-OfficeCsv` write delimited data with explicit delimiter, quoting, encoding, and spreadsheet formula-injection choices (`-FormulaInjectionPolicy`), and `Import-OfficeCsv` reads it back with typed options. Use them when the CSV is a contract with another system. Use the Excel commands when formatting, formulas, or charts are part of the result. The [safe CSV export recipe](https://github.com/EvotecIT/OfficeIMO/blob/945cb285f954b745c0f4ebfce0849c0726a76fa1/PowerShell/PSWriteOffice/Examples/Csv/Recipe-Csv-SafeExport.ps1) shows the choices side by side.

## Runnable examples

- [Edit, merge, and read YAML configuration](https://github.com/EvotecIT/OfficeIMO/blob/945cb285f954b745c0f4ebfce0849c0726a76fa1/PowerShell/PSWriteOffice/Examples/Yaml/Example-YamlEditMergeConvert.ps1)
- [TOML lock files and data conversions](https://github.com/EvotecIT/OfficeIMO/blob/945cb285f954b745c0f4ebfce0849c0726a76fa1/PowerShell/PSWriteOffice/Examples/Data/Example-DataConversions.ps1)
- [Safe CSV export](https://github.com/EvotecIT/OfficeIMO/blob/945cb285f954b745c0f4ebfce0849c0726a76fa1/PowerShell/PSWriteOffice/Examples/Csv/Recipe-Csv-SafeExport.ps1)

For every parameter, search the [command reference](/api/powershell/) for `OfficeYaml`, `OfficeToml`, `OfficeCsv`, or `Convert-OfficeDocument`. The .NET packages behind these commands are described in the [YAML](/docs/yaml/), [TOML](/docs/toml/), and [CSV](/docs/csv/) guides.
