# API map

The API map records how the public PowerShell and C# API changes in the 3.6 redesign described in the [API design](../../Docs/officeimo.api-design.md). It is the source for the 3.6 section of `MIGRATION.md`, the replacement test, and the removed-command hints in PSWriteOffice. Edit the map, not the generated outputs.

- `powershell/{area}.json`: target commands and the old-to-new mapping for each PSWriteOffice area.
- `csharp/{package}.json`: type and member changes for each C# package.

## PowerShell file

```json
{
  "area": "excel",
  "commands": [
    {
      "name": "Import-OfficeExcel",
      "alias": null,
      "status": "existing",
      "synopsis": "Reads rows from a worksheet, range or table as objects.",
      "outputType": "PSCustomObject",
      "parameterSets": [
        {
          "name": "Path",
          "default": true,
          "parameters": [
            { "name": "Path", "type": "string[]", "mandatory": true, "position": 0, "pipeline": "none", "wildcards": true },
            { "name": "Sheet", "type": "OfficeSheetSelector", "mandatory": false }
          ]
        }
      ],
      "notes": "Optional implementation notes."
    }
  ],
  "map": [
    {
      "old": "Get-OfficeExcelData",
      "kind": "cmdlet",
      "action": "merge",
      "new": "Import-OfficeExcel",
      "parameterSet": "Path",
      "parameters": { "SheetIndex": "Sheet", "SheetName": "Sheet" },
      "hint": "Import-OfficeExcel -Path .\\Book.xlsx -Range A1:D20",
      "reason": "Four commands read rows; one command with -Sheet, -Range and -Table covers them."
    }
  ]
}
```

`commands` lists every command the area has after the redesign, with its complete parameter sets. Each command also carries `examples`: a list of `{ "title", "description", "code" }` objects, at least two per command (the smallest useful call and a realistic job). Help pages are generated from them, and the validator parses each one and checks every OfficeIMO command and parameter it uses against the target surface. Common parameters (`-Verbose`, `-WhatIf`, …) are not listed; a command that supports `-WhatIf` says so with `"shouldProcess": true`.

Each `map` entry covers one current command or alias:

| Field | Meaning |
| --- | --- |
| `old` | Current command or alias name |
| `kind` | `cmdlet` or `alias` |
| `action` | `keep` (same name), `rename`, `merge` (folded into another command), `parameter` (became a parameter or parameter set of another command), `cross-format` (moved to a `Verb-OfficeDocument*` command), or `remove` |
| `new` | The replacing command or alias; `null` only for `remove` |
| `parameterSet` | The parameter set of `new` that replaces it, when relevant. Alternatives are written `A | B`; a parenthesized note after the name is allowed, as in `Path (-To Pdf)` |
| `parameters` | Renamed, moved or removed parameters: `"Old": "New"`; `"Old": "Other-OfficeCommand -Parameter"` when it moved to another command; `"Old": "Options.Property"` when it moved into the `-Options` object; `"Old": null` when removed |
| `hint` | One line shown when a script calls a removed or renamed command |
| `reason` | Why it changed. Required for every action except `keep` |

Every current command and alias appears exactly once in `map`. A retired alias may point to a command instead of a new alias.

Run `Test-ApiMap.ps1` after every change. It checks coverage, targets, parameter sets, moved parameters, approved verbs, alias rules and banned parameter names.

`remove` means the capability is gone and needs a reason that holds up to review; duplicates are `merge`, not `remove`.

## C# file

```json
{
  "package": "OfficeIMO.Word",
  "changes": [
    {
      "old": "OfficeIMO.Word.WordDocument.AsFluent()",
      "kind": "method",
      "action": "remove",
      "new": null,
      "replacement": "Use the WordDocument object model directly.",
      "reason": "One authoring model per format."
    }
  ],
  "packageMoves": [
    { "from": "OfficeIMO.Word", "to": "OfficeIMO.Word.Pdf", "scope": "Layout/ (PDF export)", "reason": "Format packages do not carry other format engines." }
  ]
}
```

`kind` is `type`, `method`, `property`, `constructor`, `enum`, `namespace` or `package`. `action` is `rename`, `remove`, `move`, `merge`, `change` (signature or return type) or `internalize`.
