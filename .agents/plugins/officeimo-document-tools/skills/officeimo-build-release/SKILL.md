---
name: officeimo-build-release
description: Use when building, packaging, releasing, or validating OfficeIMO and PSWriteOffice artifacts. Prefer repo-standard build and PowerForge/PSPublishModule entrypoints over consumer-local packaging logic.
---

# OfficeIMO Build Release

Use this skill for OfficeIMO build, package, release, website, and downstream validation work.

## Golden Path

1. Inspect the current branch and dirty state first.
2. Prefer repo-standard entrypoints.
   - Use solution and project builds for local compile evidence.
   - Use `Website/build.ps1` and website pipeline files for OfficeIMO.com changes.
   - Use shared PowerForge or PSPublishModule behavior when packaging logic is involved.
3. Keep generated output narrow.
   - If docs are generated, update the source of truth and regenerate.
   - If a validation command dirties broad generated docs outside the intended change, inspect and keep only intentional files.
4. Distinguish release layers before answering whether a fix is available.
   - Local source
   - Open PR
   - Merged branch
   - Release tag
   - NuGet or PSGallery package
   - Installed downstream version
5. For PDF or conversion release work, include at least one artifact-level proof, not just compile success.

## Useful Checks

```powershell
dotnet build OfficeIMO.sln -c Release
dotnet test OfficeIMO.Word.Tests\OfficeIMO.Word.Tests.csproj -c Release
dotnet test OfficeIMO.Excel.Tests\OfficeIMO.Excel.Tests.csproj -c Release
dotnet test OfficeIMO.Shared.Tests\OfficeIMO.Shared.Tests.csproj -c Release
pwsh Website\build.ps1 -Mode dev
git diff --check
```

Run the smallest useful subset when the task is narrow, and expand when public contracts, packaging, or generated website output changes.
