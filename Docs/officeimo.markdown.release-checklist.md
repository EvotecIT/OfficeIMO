# OfficeIMO Markdown Release Checklist

Use this checklist when publishing a new `OfficeIMO.Markdown` / `OfficeIMO.MarkdownRenderer` package line.

## Goal

Ship a deliberate markdown package baseline that:

- matches the current public docs
- has green solution/build coverage on the supported targets
- has a benchmark harness available for quick sanity checks
- is ready for immediate downstream adoption in `IntelligenceX`

## Before Bumping Versions

- Confirm the markdown-specific release-prep PRs are merged:
  - public naming/docs cleanup
  - benchmark harness
  - root/package README polish
- Confirm `OfficeIMO.Markdown` and `OfficeIMO.MarkdownRenderer` still build cleanly from `master`.
- Confirm no markdown-specific PRs remain open except intentionally deferred follow-up work.

## Validation Baseline

Run at minimum:

```powershell
dotnet build .\OfficeIMO.Markdown\OfficeIMO.Markdown.csproj -c Release -f netstandard2.0
dotnet build .\OfficeIMO.MarkdownRenderer\OfficeIMO.MarkdownRenderer.csproj -c Release
dotnet test .\OfficeIMO.Tests\OfficeIMO.Tests.csproj --framework net8.0 --filter "FullyQualifiedName~Markdown_"
dotnet run -c Release --project .\OfficeIMO.Markdown.Benchmarks\OfficeIMO.Markdown.Benchmarks.csproj -- --list flat
```

Recommended additional validation:

- run one focused benchmark class to confirm the harness still executes
- recheck `OfficeIMO.Markdown\README.md` and `OfficeIMO.Markdown\COMPATIBILITY.md`
- recheck `OfficeIMO.MarkdownRenderer\README.md` for preset/composition and visual-contract guidance
- verify package descriptions/readmes/icons render correctly in the `.csproj` metadata
- run `dotnet pack` for both markdown packages and confirm the expected versioned `.nupkg` artifacts are produced

## Publish Sequence

1. Update `VersionPrefix` in:
   - `OfficeIMO.Markdown\OfficeIMO.Markdown.csproj`
   - `OfficeIMO.MarkdownRenderer\OfficeIMO.MarkdownRenderer.csproj`
2. Update any release notes/changelog entries that mention the new markdown baseline.
3. Publish `OfficeIMO.Markdown`.
4. Publish `OfficeIMO.MarkdownRenderer`.
5. Verify the packages are visible on NuGet with the expected README/description metadata.

## Immediate Downstream Follow-up

After publish, open the fresh `IntelligenceX` package-adoption PR and:

1. update `IntelligenceX.Chat\Directory.Build.props`
2. rerun the merged OfficeIMO package-mode gate
3. revalidate UI render, markdown export, and DOCX export
4. merge the adoption PR only after those checks are green

## Non-Goals For This Release

Do not wait for:

- full CommonMark/GFM conformance
- a fully finished extension ecosystem
- every future markdown feature idea

The goal is a stable, intentional baseline that downstream code can consume and validate.
