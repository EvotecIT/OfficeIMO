# OfficeIMO.Tool

One command-line entry point for OfficeIMO document workflows:

```powershell
dotnet tool install --global OfficeIMO.Tool

officeimo html capabilities
officeimo reader read document.docx --format markdown
officeimo markup validate document.markup --profile document
```

Commands are grouped by capability so their contracts remain explicit:

- `officeimo html` converts HTML or MHTML to PDF and reports renderer capabilities.
- `officeimo reader` extracts supported documents as Markdown or JSON.
- `officeimo markup` parses, validates, emits, previews, and exports OfficeIMO Markup.

Run `officeimo help` or `officeimo <area> --help` for the complete command contract.

## Exit codes

| Code | Meaning |
| ---: | --- |
| `0` | Success |
| `1` | The requested validation completed and found document errors |
| `2` | Invalid command or option |
| `3` | Input was not found |
| `4` | Input is unsupported or an I/O operation failed |
| `5` | The document operation failed |
| `6` | Output failed or conversion completed with error-severity diagnostics |
| `130` | Cancelled |

## Reader tool migration

`OfficeIMO.Reader.Tool` 3.0.0 used the `officeimo-reader` executable. New releases use
the unified package and add the `reader` command area:

```powershell
# Before
officeimo-reader read document.docx --format markdown

# Now
officeimo reader read document.docx --format markdown
```

The Reader command contract remains explicit; the migration does not add a compatibility
shim or duplicate Reader implementation.
