# OfficeIMO Security benchmarks

This opt-in project compares equivalent detached CMS signing and
verification through `OfficeIMO.Security` and .NET
`SignedCms`. Both lanes use the same RSA-2048 certificate, SHA-256 digest,
embedded signer certificate, and detached content. Signing time, chain
validation, revocation, timestamps, downloads, and external trust work are
disabled in both lanes.

The verification comparison performs signature-only cryptographic checking,
extracts the signer certificate and core signer metadata, and evaluates the
digital-signature key usage in both lanes. OfficeIMO additionally materializes
its typed policy/findings model; interpret any remaining small-payload fixed
allocation difference in that context.

Preflight cross-verifies both produced signatures through both engines,
requires one signer and one embedded certificate, checks the detached and
SHA-256 contracts, and requires both engines to reject tampered content.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\security\evidence.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- --filter '*SecurityCms*' --job Short --noOverwrite
```

The isolated evidence command records per-operation elapsed time and managed
allocation plus retained managed heap, sampled managed peak, absolute process
peak working set, CMS/content sizes, runtime, OS, commit, and dirty-tree state.

The project stays outside `OfficeIMO.sln`. BenchmarkDotNet and the platform
PKCS package remain benchmark-only dependencies and do not enter OfficeIMO
runtime projects or packages.
