# Access verification

This project is an executable public API consumer and an opt-in Windows producer/oracle route. It has no production dependency on Access or DAO. `New-AccessCorpus.ps1` and `New-AccessApplicationCorpus.ps1` require an installed Windows engine and always create a fresh output directory. They never open user databases or replace existing output.

```powershell
$scratchRoot = $env:EVOTEC_SCRATCH_ROOT
if (-not $scratchRoot) { throw 'Configure and verify a development scratch volume first.' }
./OfficeIMO.Access.Verification/New-AccessCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-table-corpus')
./OfficeIMO.Access.Verification/New-AccessApplicationCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-application-corpus')
./OfficeIMO.Access.Verification/Test-NativeBootstrap.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-bootstrap-control')
```

The bootstrap command is a failed-creation negative control; its rejection is not writer qualification. The manifest names the next real native bootstrap experiment. The application corpus does not execute macros or modules, use user databases, change trust policy, or register a database provider. It exports a simple report to PDF through Access and releases its owned application. Macro, protection, signatures and newer generation fixtures remain explicit evidence gaps.

Run the consumer and generate the operation contract from the canonical library catalog:

```sh
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- OfficeIMO.Access.Tests/Fixtures
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- --catalog OfficeIMO.Access
dotnet test OfficeIMO.Access.Tests/OfficeIMO.Access.Tests.csproj -c Release
```

The generated JSON and Markdown describe implemented behavior and unsupported operations. Human-authored README and roadmap wording is not pinned by product unit tests. The verification executable stays outside the normal solution; the library and deterministic Access contract tests are in the solution and established test workflow.
