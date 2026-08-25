# OfficeIMO.Security detached CMS evidence (2026-08-24)

## Result

OfficeIMO.Security's detached RSA CMS signing and verification paths now sit
close to or ahead of the .NET platform implementation for equivalent work. The
pre-optimization verification path was an incident: depending on content size
and signature producer, it took 24.6-171.5x the platform time and allocated a
nearly fixed 0.95 MiB per operation. The optimized path is 0.68-1.87x platform
time and 0.025-1.79x platform allocation across the controlled matrix. Every
measured time, allocation, retained-heap, sampled managed-peak, and process-peak
ratio is below the 2x contender boundary. This is a tolerable margin, not a
claim that the remaining 1.68-1.87x small-workload gaps are optimal.

## Equivalent contract

Both engines use the same deterministic 1 KiB, 64 KiB, and 1 MiB content, the
same ephemeral RSA-2048 self-signed certificate, SHA-256, detached content, one
embedded end certificate, and no signing time. Verification performs
signature-only cryptographic checking, materializes the signer certificate and
core signer metadata, inspects standard signed attributes, and evaluates
digital-signature key usage. Chain building, revocation, certificate downloads,
timestamp validation, and external trust are disabled in both lanes.

Preflight cross-verifies both generated signatures through both engines,
requires one signer and one embedded certificate, checks the detached and
SHA-256 contracts, and requires rejection of tampered content. OfficeIMO also
returns its typed policy, certificate-validation, and findings model, which is
part of its public result contract.

The comparison uses .NET `SignedCms`, the platform API for CMS/PKCS #7. On
.NET 8 and later OfficeIMO also uses this platform primitive for the common
attribute-free, one-signer detached RSA/SHA shape, then materializes the same
typed OfficeIMO policy result. Signed-attribute, timestamp, chain-validation,
multi-signer, non-data, and non-RSA structures remain on the complete Bouncy
Castle verifier. `netstandard2.0` and .NET Framework retain that complete path.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. Ratios are OfficeIMO divided
by platform, so lower is better.

### Signing

| Content | OfficeIMO mean | Platform mean | Time ratio | OfficeIMO allocation | Platform allocation | Allocation ratio | CMS size ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 KiB | 389.7 us | 448.6 us | 0.87x | 25.79 KiB | 25.05 KiB | 1.03x | 1.105x |
| 64 KiB | 451.8 us | 480.7 us | 0.94x | 25.79 KiB | 88.05 KiB | 0.29x | 1.105x |
| 1 MiB | 826.0 us | 941.3 us | 0.88x | 25.77 KiB | 1,049.07 KiB | 0.025x | 1.105x |

OfficeIMO includes the CMS algorithm-protection signed attribute and the
platform output does not, accounting for roughly 124 additional bytes. The
1.105x size ratio is therefore a documented integrity-policy cost rather than
unexplained package growth.

### Verification

| Signature producer | Content | OfficeIMO mean | Platform mean | Time ratio | OfficeIMO allocation | Platform allocation | Allocation ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| OfficeIMO | 1 KiB | 189.30 us | 264.13 us | 0.72x | 37.24 KiB | 20.78 KiB | 1.79x |
| OfficeIMO | 64 KiB | 211.46 us | 310.01 us | 0.68x | 37.24 KiB | 84.03 KiB | 0.44x |
| OfficeIMO | 1 MiB | 613.48 us | 795.56 us | 0.77x | 37.21 KiB | 1,044.92 KiB | 0.036x |
| Platform | 1 KiB | 189.71 us | 101.48 us | 1.87x | 15.27 KiB | 14.33 KiB | 1.07x |
| Platform | 64 KiB | 224.75 us | 133.68 us | 1.68x | 78.40 KiB | 77.46 KiB | 1.01x |
| Platform | 1 MiB | 722.74 us | 587.33 us | 1.23x | 1,039.14 KiB | 1,038.04 KiB | 1.00x |

The optimized verifier uses the .NET CMS parser for the narrow attribute-free
shape it handles completely. Richer structures are preflighted without first
allocating a second parser, then use one Bouncy Castle parse, platform SHA and
RSA primitives for SHA-1/SHA-256/SHA-384/SHA-512 PKCS #1 signatures, reused
certificate encodings, fixed result arrays, and lazy finding collections.

## Isolated memory and provenance evidence

The checked-in evidence runner starts a fresh child process for each engine,
operation, scale, producer, and repetition. The table below reports the median
OfficeIMO/platform ratio over three repetitions from commit
`7cbd8e351de8919e4ba8d5a1f786df0fc7869450`; the runner recorded a clean source
tree on .NET 10.0.11, Windows build 26200, x64, with 32 logical processors.

| Operation | Producer | Content | Time | Allocation | Retained heap | Managed peak | Process peak |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| Sign | OfficeIMO/platform | 1 KiB | 0.99x | 1.04x | 1.15x | 1.04x | 1.00x |
| Sign | OfficeIMO/platform | 64 KiB | 0.91x | 0.30x | 1.15x | 0.30x | 0.97x |
| Sign | OfficeIMO/platform | 1 MiB | 0.54x | 0.025x | 1.08x | 0.11x | 0.94x |
| Verify | OfficeIMO signature | 1 KiB | 0.79x | 1.69x | 0.061x | 1.68x | 1.02x |
| Verify | OfficeIMO signature | 64 KiB | 0.75x | 0.44x | 0.12x | 0.45x | 0.97x |
| Verify | OfficeIMO signature | 1 MiB | 0.93x | 0.036x | 0.86x | 0.15x | 0.94x |
| Verify | Platform signature | 1 KiB | 1.64x | 1.09x | 1.49x | 1.08x | 1.00x |
| Verify | Platform signature | 64 KiB | 1.48x | 1.02x | 1.47x | 1.01x | 1.00x |
| Verify | Platform signature | 1 MiB | 1.24x | 1.00x | 1.18x | 1.00x | 1.01x |

Absolute median process peaks are 39.6-47.4 MiB across all lanes. Retained
managed growth is no more than 23.46 KiB for OfficeIMO and no more than 15.80
KiB for the platform verifier. CMS artifacts are 1,304-1,306 bytes for
OfficeIMO and 1,180-1,182 bytes for the platform implementation.

## Reproduce

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- --filter '*SecurityCms*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Security.Benchmarks -- evidence --repeat 3 --json .benchmark-artifacts\security\evidence.json
```

Raw BenchmarkDotNet output and process evidence remain ignored machine-local
artifacts. This note retains the compact reproducible result and the exact
source provenance.
