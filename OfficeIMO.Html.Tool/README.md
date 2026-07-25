# OfficeIMO.Html.Tool

`OfficeIMO.Html.Tool` converts bounded HTML or MHTML input to PDF through the same renderer used by the .NET API. It has no browser process, JavaScript engine, network client, or third-party command parser.

## Install

```powershell
dotnet tool install --global OfficeIMO.Html.Tool
```

## Convert

```powershell
officeimo-html convert report.html
officeimo-html convert archive.mhtml --output report.pdf
Get-Content report.html -Raw | officeimo-html convert - --input-format html --output report.pdf
officeimo-html convert report.html --stylesheet print.css --max-pages 500
```

The default policy accepts bounded data URIs and resources embedded in MHTML. It does not read local or remote resources. `--stylesheet` is repeatable up to 16 times and each stylesheet is limited to 4 MiB.

`--pdf-ua-language en-US` enables PDF/UA-1 groundwork and reports internal readiness. It deliberately does not claim conformance without passing external validator evidence.

Use `--force` to replace an existing output. File output is written to a temporary sibling and moved into place only after serialization succeeds.

## Inspect capabilities

```powershell
officeimo-html capabilities
officeimo-html capabilities --format json
```

## Exit codes

| Code | Meaning |
| ---: | --- |
| 0 | Success |
| 2 | Invalid command or arguments |
| 3 | Input file not found |
| 4 | Input/output failure |
| 5 | Conversion failure |
| 6 | PDF created with error-severity conversion diagnostics |
| 130 | Cancelled |
For a deterministic PDF/UA-ready artifact, provide an embedded TrueType family:

```powershell
officeimo-html convert report.html --output report.pdf `
  --pdf-ua-language en-US `
  --font-family "Document Sans" `
  --font-regular .\fonts\DocumentSans-Regular.ttf `
  --font-bold .\fonts\DocumentSans-Bold.ttf
```

The tool reports internal readiness separately from independent validator evidence.
