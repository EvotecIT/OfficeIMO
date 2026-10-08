# Access verification

This project is an executable public API consumer and an opt-in Windows producer/oracle route. It has no production dependency on Access or DAO. `New-AccessCorpus.ps1` and `New-AccessApplicationCorpus.ps1` require an installed Windows engine and always create a fresh output directory. They never open user databases or replace existing output.

```powershell
$scratchRoot = $env:EVOTEC_SCRATCH_ROOT
if (-not $scratchRoot) { throw 'Configure and verify a development scratch volume first.' }
./OfficeIMO.Access.Verification/New-AccessCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-table-corpus')
./OfficeIMO.Access.Verification/New-AccessApplicationCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-application-corpus')
./OfficeIMO.Access.Verification/New-AccessProfileCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-profile-corpus')
./OfficeIMO.Access.Verification/Test-NativeBootstrap.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-bootstrap-control')
./OfficeIMO.Access.Verification/Test-NativeCreation.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-native-creation') -ArtifactsDirectory (Join-Path $scratchRoot 'access-verification-build')
```

`Test-NativeBootstrap` is a failed-creation negative control; its rejection is not writer qualification. `Test-NativeCreation` generates MDB and ACCDB pages from logical definitions without seeds, then independently observes tables, typed rows, primary/foreign indexes, index seek and relationships through DAO. Constraint violations are tested on separate owned copies. The resulting manifest and expected-value exports qualify a fixed feasibility spike; the production package still has no native Save codec. The generator can also run portably through `--bootstrap <fresh-directory>`; DAO consumption requires Windows.

The application corpus creates and exports a simple form/report, inert module and `StopMacro`, exports the report to PDF, and releases its owned application. The profile corpus creates only synthetic protection inputs, BigInt values and an extended date/time field definition; it records the unavailable Jet 3 producer. Signature, native application-carrier, complex-field and protection-codec qualification remain explicit later-milestone work.

Run the consumer and generate the operation contract from the canonical library catalog:

```sh
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- OfficeIMO.Access.Tests/Fixtures
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- --catalog OfficeIMO.Access
dotnet test OfficeIMO.Access.Tests/OfficeIMO.Access.Tests.csproj -c Release
```

The generated JSON and Markdown describe implemented behavior and unsupported operations. Human-authored README and roadmap wording is not pinned by product unit tests. The verification executable stays outside the normal solution; the library and deterministic Access contract tests are in the solution and established test workflow.
