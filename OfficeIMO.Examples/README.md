# OfficeIMO.Examples — Runnable Samples

This project contains small, focused samples that demonstrate usage across Word, Excel, PowerPoint, PDF, and Visio packages.

- Build and run to explore features end-to-end.
- Output is designed to be human-readable and PowerShell-friendly where applicable.
- Run `dotnet run --project OfficeIMO.Examples -- --pdf-professional` to generate a professional `OfficeIMO.Pdf` report sample, `--pdf-table-styles` for the Word-like table style gallery, `--pdf-showcase` for richer statement/dashboard/manipulation PDFs, or `--pdf` for the full first-party PDF example set.
- PowerPoint examples include direct slide building plus designer decks, design-brief recommendations, and semantic deck-plan scoring.
- Run `dotnet run --project OfficeIMO.Examples --framework net8.0 -- --visio-showcase` to generate a curated Visio set from basic fluent shapes through advanced flowchart, block, swimlane, sequence, network, architecture, editing, and quality-gallery diagrams.
- Add `--visio-export` when Microsoft Visio desktop is installed to export first-page `.png` and `.svg` previews plus a browseable `index.html` into `Documents/Visio Showcase/Preview`.

