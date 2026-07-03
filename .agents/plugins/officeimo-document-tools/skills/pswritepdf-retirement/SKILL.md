---
name: pswritepdf-retirement
description: Use when retiring PSWritePDF behavior into OfficeIMO and PSWriteOffice, comparing PDF gaps, or deciding where PowerShell-facing PDF features should live. Keep reusable PDF behavior in OfficeIMO and PSWriteOffice cmdlets thin and friendly.
---

# PSWritePDF Retirement

Use this skill for PSWritePDF retirement and PSWriteOffice PDF surface work.

## Golden Path

1. Read the current retirement gap analysis first:
   `Roadmap/PSWritePDF-RetirementGapAnalysis.md`
2. Inventory the OfficeIMO reusable capability before adding PowerShell code.
3. Put PDF generation, parsing, metadata, conversion, and rendering behavior in OfficeIMO or the shared PDF engine.
4. Keep PSWriteOffice as a friendly command surface.
   - Parameters should map to reusable options.
   - Presets and themes are welcome when they hide complexity without hiding behavior.
   - Avoid cmdlet-specific forks of conversion or PDF logic.
5. Validate both developer-source and package-style usage when a change crosses repo boundaries.

## Questions To Answer Before Coding

- Is this a core PDF capability, a conversion adapter, or a PowerShell UX feature?
- Does OfficeIMO already have the right byte or stream API?
- Will Blazor WebAssembly need this behavior too?
- Is the downstream PSWriteOffice package using local project references, a packed artifact, or a published NuGet version?

## Useful Checks

```powershell
dotnet build OfficeIMO.Word.Pdf\OfficeIMO.Word.Pdf.csproj -c Release -f net10.0
dotnet test OfficeIMO.Pdf.Tests\OfficeIMO.Pdf.Tests.csproj -c Release --filter "*Pdf*"
```

For PSWriteOffice changes, also validate import and command discovery from the target PSWriteOffice worktree.
