# OfficeIMO.Examples — Runnable Samples

This project contains small, focused samples that demonstrate usage across Word, Excel, PowerPoint, PDF, and Visio packages.

- Build and run to explore features end-to-end.
- Output is designed to be human-readable and PowerShell-friendly where applicable.
- Run `dotnet run --project OfficeIMO.Examples -- --pdf-professional` to generate a professional `OfficeIMO.Pdf` report sample, `--pdf-table-styles` for the Word-like table style gallery, `--pdf-showcase` for richer statement/dashboard/manipulation PDFs, or `--pdf` for the full first-party PDF example set.
- PowerPoint examples include direct slide building plus designer decks, design-brief recommendations, and semantic deck-plan scoring.
- Run `dotnet run --project OfficeIMO.Examples --framework net8.0 -- --visio-showcase` to generate a curated Visio set from basic fluent shapes through advanced flowchart, block, swimlane, sequence, network, architecture, editing, data-driven inventory, and quality-gallery diagrams.
- Add `--visio-export` when Microsoft Visio desktop is installed to export first-page `.png` and `.svg` previews plus a browseable `index.html` into `Documents/Visio Showcase/Preview`.
- Run `dotnet run --project OfficeIMO.Examples --framework net8.0 -- --visio-integration-stencils C:\StencilPacks\Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio` to generate a real graph from an external multi-`.vssx` Microsoft Integration/Azure stencil pack. The same example can be included in the Visio showcase by setting `OFFICEIMO_VISIO_INTEGRATION_STENCILS` to that pack root before running `--visio-showcase`.
- Run `dotnet run --project OfficeIMO.Examples --framework net8.0 -- --visio-stencil-gallery C:\StencilPacks\Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio` to generate a contact-sheet VSDX that shows which masters OfficeIMO can load from an external stencil pack or pack directory.
- Stencil catalogs can resolve friendly fallback queries with `catalog.FindBest("API Management Services", "Application Gateway")`, and graph diagrams can place catalog-backed nodes directly with `graph.StencilNode("gateway", "API gateway", catalog, "API Management Services", "Application Gateway")`.
