# Access verification

This project is an executable public API consumer and an opt-in Windows producer/oracle route. It has no production dependency on Access or DAO. `New-AccessCorpus.ps1` and `New-AccessApplicationCorpus.ps1` require an installed Windows engine and always create a fresh output directory. They never open user databases or replace existing output.

```powershell
$scratchRoot = $env:EVOTEC_SCRATCH_ROOT
if (-not $scratchRoot) { throw 'Configure and verify a development scratch volume first.' }
./OfficeIMO.Access.Verification/New-AccessCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-table-corpus')
./OfficeIMO.Access.Verification/New-AccessApplicationCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-application-corpus')
./OfficeIMO.Access.Verification/New-AccessProfileCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-profile-corpus')
./OfficeIMO.Access.Verification/New-AccessReaderCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-reader-corpus')
./OfficeIMO.Access.Verification/New-AccessGenerationCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-generation-corpus')
./OfficeIMO.Access.Verification/New-AccessDesignerCorpus.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-designer-corpus') -IncludeDataMacro -IncludeEmbeddedMacro
./OfficeIMO.Access.Verification/Test-NativeBootstrap.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-bootstrap-control')
./OfficeIMO.Access.Verification/Test-NativeCreation.ps1 -OutputDirectory (Join-Path $scratchRoot 'access-native-creation') -ArtifactsDirectory (Join-Path $scratchRoot 'access-verification-build')
```

`Test-NativeBootstrap` is a failed-creation negative control; its rejection is not writer qualification. `Test-NativeCreation` runs the production `Create`/`Save` API without seeds, then independently checks MDB/ACCDB tables, scalar/long values, AutoNumber seeds, primary/text/composite index seeks and relationships through DAO. Boundary cases include 255 ordinary and long-value columns and allocation beyond 512 pages. Separate owned Access instances disable macros, reopen copies, reject constraint violations, update/append/delete rows and save. OfficeIMO independently reads those saved changes. Compact JSON observations record the files and hashes. The portable generation route is `--create <fresh-directory>`; independent Access consumption requires Windows.

The application corpus creates and exports a simple form/report, inert module and `StopMacro`, exports the report to PDF, and releases its owned application. The designer corpus adds bound sources, sections/controls, CP1250 VBA source, ACE embedded/data macros and independent text exports. The profile corpus creates only synthetic protection inputs, BigInt values and an extended date/time definition; it records the unavailable Jet 3 producer. Signature/protection operations remain separate qualification work.

The reader corpus qualifies typed scalars, deleted/grown rows, composite relationships/indexes, rich text, lookup properties, multivalued text and attachment metadata/bytes. It uses the installed ACE OLE DB engine only to define Decimal precision/scale, then independently observes persisted data through read-only DAO. The generation corpus uses Access format 9/10 MDB creation and read-only DAO observations for legacy values and saved queries. It also qualifies calculated-field opaque representations, Large Number and extended dates. A controlled scale-seven extended-date value mutation is explicitly labeled in its manifest and independently consumed by DAO; the host's producer persists scale zero. Neither producer route enters the shipped library or its dependency graph.

The executable consumer exercises the public native API, `DataTable.Load`, selective schema loading, incremental binary access, structured values, redacted inert links, precision-sensitive values and unsupported query authoring/conversion. The checked-in fixture manifests own hashes, byte sizes, provenance and observations.

To prove unchanged application preservation, run `--preserve OfficeIMO.Access.Tests/Fixtures <fresh-directory>`, then `Test-NativePreservation.ps1 -OutputDirectory <fresh-directory>`. The managed route checks whole-file bytes and application-stream hashes; Access independently reopens and re-exports the preserved objects. All application content remains inert.

Run the consumer and generate the operation contract from the canonical library catalog:

```sh
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- OfficeIMO.Access.Tests/Fixtures
dotnet run --project OfficeIMO.Access.Verification/OfficeIMO.Access.Verification.csproj -- --catalog OfficeIMO.Access
dotnet test OfficeIMO.Access.Tests/OfficeIMO.Access.Tests.csproj -c Release
```

The generated JSON and Markdown describe implemented behavior and unsupported operations. Human-authored README and roadmap wording is not pinned by product unit tests. The verification executable stays outside the normal solution; the library and deterministic Access contract tests are in the solution and established test workflow.
