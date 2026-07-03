# OfficeIMO Legacy DOC Corpus

This folder is for small, reviewable `.doc` fixtures that protect legacy binary
Word import behavior.

For each `sample.doc`, keep an approved `sample.import-report.md` generated from
`LegacyDocImportReport.ToMarkdown()`. The corpus test compares the current import
report against the approved baseline so parser, projection, diagnostics, and
preserve-only signals cannot drift silently.

To refresh approved baselines after an intentional import change:

```powershell
$env:OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES = '1'
dotnet test .\OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj --filter "FullyQualifiedName~LegacyDoc_CorpusImportReports_MatchCheckedInBaselines"
Remove-Item Env:\OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES
```

Keep fixtures focused and document their source or generator in a short note next
to the document when possible. Do not include sensitive customer data.

## Optional Word COM Fixture Generation

When Microsoft Word is installed on Windows, use the local generator to create
repeatable Word 97-2003 `.doc` samples for simple paragraph, character
formatting, and paragraph formatting coverage:

```powershell
.\OfficeIMO.TestAssets\Documents\LegacyDocCorpus\Generate-WordComFixtures.ps1
```

The generator skips existing files by default. Use `-Force` only when you intend
to refresh checked-in fixtures and their matching import reports:

```powershell
.\OfficeIMO.TestAssets\Documents\LegacyDocCorpus\Generate-WordComFixtures.ps1 -Force
$env:OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES = '1'
dotnet test .\OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj --filter "FullyQualifiedName~LegacyDoc_CorpusImportReports_MatchCheckedInBaselines"
Remove-Item Env:\OFFICEIMO_UPDATE_LEGACY_DOC_CORPUS_BASELINES
```

Use `-Scenario` to generate one focused fixture family at a time:

```powershell
.\OfficeIMO.TestAssets\Documents\LegacyDocCorpus\Generate-WordComFixtures.ps1 -Scenario CharacterFormatting
```

The generator is test-support tooling only. It must not be referenced by
`OfficeIMO.Word` or any production project.

## Optional Desktop Word Validation

When Microsoft Word is installed on Windows, run the opt-in COM lane to generate
a real `.doc` document through Word, import it through OfficeIMO, save it back
through the native `.doc` writer, and verify desktop Word opens both files:

```powershell
$env:OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION = '1'
dotnet test .\OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj --filter "FullyQualifiedName~LegacyDoc_ComGeneratedDocument_ImportsAndNativeSaveOpensInDesktopWordWhenRequested"
Remove-Item Env:\OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION
```

The same switch also verifies checked-in legacy DOC corpus fixtures open through
desktop Word:

```powershell
$env:OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION = '1'
dotnet test .\OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj --filter "FullyQualifiedName~LegacyDoc_CorpusFixtures_OpenInDesktopWordWhenRequested"
Remove-Item Env:\OFFICEIMO_RUN_LEGACY_DOC_COM_VALIDATION
```
