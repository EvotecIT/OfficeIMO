# Independent Project file verification

This opt-in console tool reads OfficeIMO-generated files with MPXJ.Net and exports MSPDI, MPX, or JSON for independent inspection. It is outside the normal solution, is not packable, and is never referenced by the OfficeIMO runtime projects. Building it restores the MPXJ/IKVM dependency tree; normal OfficeIMO builds do not need it.

## Run a comparison

Run from the repository root and choose new output paths:

```powershell
dotnet build OfficeIMO.Project.Verification -c Release
dotnet build OfficeIMO.Project.IndependentVerification -c Release

dotnet run --project OfficeIMO.Project.Verification -c Release --no-build -- `
    conversion-matrix ./artifacts/project/conversions

dotnet run --project OfficeIMO.Project.IndependentVerification -c Release --no-build -- `
    batch-mspdi ./artifacts/project/conversions '*-to-*'

dotnet run --project OfficeIMO.Project.Verification -c Release --no-build -- `
    native-corpus ./artifacts/project/conversions ./artifacts/project/comparison '*-to-*'
```

`batch-mspdi` exports each MPP, MPT, or MPX match to a sibling `.xml` and writes `independent-batch.json`. It refuses existing output paths. The comparison records unexplained differences as failures and keeps qualified defaults, cache differences, and derived values as explicit observations.

For one file, use `<input> <new-output> <MSPDI|MPX|JSON>`. For example:

```powershell
dotnet run --project OfficeIMO.Project.IndependentVerification -c Release --no-build -- `
    ./artifacts/project/authored.mpp ./artifacts/project/authored.xml MSPDI
```

## Interpret the evidence

An independent reader accepting a file does not prove that Microsoft Project will open it or preserve every stored value. Run the [application oracle](../Build/Project/README.md) separately for supported installed formats. Some independent exports derive critical flags, remaining duration, or WBS; the comparison retains these differences instead of changing the imported model to match the oracle.

The [historical fixture manifest](../Build/Project/historical-fixtures.json) identifies selected upstream binaries by immutable URL and SHA-256. They remain external verification inputs under the upstream repository's license; no fixture binary, implementation code, or runtime dependency is copied into OfficeIMO. A listed fixture can represent an unsupported dialect: consult the package's [support matrix](../OfficeIMO.Project/SUPPORT.md) before treating it as a passing case.
