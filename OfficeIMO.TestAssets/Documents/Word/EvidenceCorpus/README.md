# Word evidence corpus provenance

This corpus is the executable index for modern and legacy Word evidence that was
previously spread across focused test folders. It does not create a blanket
compatibility claim. Each manifest entry identifies its producer, immutable
artifact hash or deterministic generator, applicable oracle, focused contract,
and required loss policy.

Binary producer artifacts use raw-byte hashes. Repository-authored text fixtures
use `canonical-text`, which normalizes line endings to LF before hashing so
Windows and POSIX checkouts prove the same content.

Microsoft Word fixtures were authored or refreshed through desktop Word and are
used as interoperability inputs, not as runtime dependencies. OfficeIMO source
cases are deterministic, repository-owned inputs whose producer is the named
OfficeIMO converter, renderer, workflow, or benchmark generator. The legacy DOC
entry remains guarded by its import report and by the normal loss preflight.

The corpus deliberately contains no claim for arbitrary native DOC authoring.
OfficeIMO supports a documented first-party DOC subset; callers must use
`AssessLegacyDocWrite`, conversion reports, and an explicit loss decision for
content outside that subset.

Run the manifest contract with:

```powershell
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net8.0 --filter "FullyQualifiedName~WordEvidenceCorpus"
```

Run the opt-in Microsoft Word layout oracle with:

```powershell
$env:OFFICEIMO_RUN_WORD_LAYOUT_COM_VALIDATION = '1'
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -f net8.0 --filter "FullyQualifiedName~DrawingLayout_AnchoredGroupMatchesDesktopWordGeometry"
```
