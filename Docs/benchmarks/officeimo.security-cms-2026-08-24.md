# OfficeIMO.Security detached CMS evidence (2026-08-24)

## Result

OfficeIMO.Security's detached RSA CMS signing and verification paths now sit
close to or ahead of the .NET platform implementation for equivalent work. The
pre-optimization verification path was an incident: depending on content size
and signature producer, it took 24.6-171.5x the platform time and allocated a
nearly fixed 0.95 MiB per operation. The optimized path is 0.29-1.09x platform
time and 0.03-2.11x platform allocation.

The smallest platform-produced signature is not yet inside the repository's
contender boundary for every memory metric. Its OfficeIMO verification uses
2.11x the platform allocation in BenchmarkDotNet; the isolated runner records
2.09x allocation and 2.07x sampled managed-heap peak. This remains open
remediation rather than being described as parity.

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

The comparison uses .NET `SignedCms`, the platform API for CMS/PKCS #7. Its
package dependency remains isolated in the opt-in benchmark project and does
not enter OfficeIMO runtime projects or normal solution restore.

## BenchmarkDotNet results

Environment: Windows 11 build 26200, AMD Ryzen 9 9950X3D2, .NET SDK 10.0.111,
.NET 10.0.11 x64, BenchmarkDotNet 0.15.8 `ShortRun`. Ratios are OfficeIMO divided
by platform, so lower is better.

### Signing

| Content | OfficeIMO mean | Platform mean | Time ratio | OfficeIMO allocation | Platform allocation | Allocation ratio | CMS size ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 KiB | 367.1 us | 526.1 us | 0.70x | 25.79 KiB | 25.05 KiB | 1.03x | 1.105x |
| 64 KiB | 464.4 us | 494.3 us | 0.94x | 25.79 KiB | 88.06 KiB | 0.29x | 1.105x |
| 1 MiB | 930.5 us | 1,065.7 us | 0.87x | 25.79 KiB | 1,049.23 KiB | 0.025x | 1.105x |

OfficeIMO includes the CMS algorithm-protection signed attribute and the
platform output does not, accounting for roughly 124 additional bytes. The
1.105x size ratio is therefore a documented integrity-policy cost rather than
unexplained package growth.

### Verification

| Signature producer | Content | OfficeIMO mean | Platform mean | Time ratio | OfficeIMO allocation | Platform allocation | Allocation ratio |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| OfficeIMO | 1 KiB | 69.83 us | 242.32 us | 0.29x | 37.67 KiB | 20.84 KiB | 1.81x |
| OfficeIMO | 64 KiB | 123.69 us | 311.35 us | 0.40x | 37.66 KiB | 83.97 KiB | 0.45x |
| OfficeIMO | 1 MiB | 583.26 us | 823.13 us | 0.71x | 38.51 KiB | 1,044.89 KiB | 0.037x |
| Platform | 1 KiB | 85.65 us | 78.44 us | 1.09x | 30.53 KiB | 14.45 KiB | **2.11x** |
| Platform | 64 KiB | 110.72 us | 153.73 us | 0.72x | 30.53 KiB | 77.64 KiB | 0.39x |
| Platform | 1 MiB | 542.74 us | 601.63 us | 0.90x | 30.56 KiB | 1,037.87 KiB | 0.029x |

The optimized verifier parses the detached CMS once for standard RSA signers,
uses platform SHA and RSA primitives for SHA-1/SHA-256/SHA-384/SHA-512 PKCS #1
signatures, reuses certificate encodings, and preserves the Bouncy Castle
fallback for other algorithms and structures.

## Isolated memory and provenance evidence

The checked-in evidence runner starts a fresh child process for each engine,
operation, scale, producer, and repetition. The table below reports the median
OfficeIMO/platform ratio over three repetitions from commit
`656a6cd1da6d072edc4de384ce5b20f24d0d000a`; the runner recorded a clean source
tree on .NET 10.0.11, Windows build 26200, x64, with 32 logical processors.

| Operation | Producer | Content | Time | Allocation | Retained heap | Managed peak | Process peak |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |
| Sign | OfficeIMO/platform | 1 KiB | 0.88x | 1.04x | 1.09x | 1.04x | 1.00x |
| Sign | OfficeIMO/platform | 64 KiB | 0.95x | 0.30x | 1.09x | 0.30x | 0.96x |
| Sign | OfficeIMO/platform | 1 MiB | 0.81x | 0.025x | 1.07x | 0.22x | 0.96x |
| Verify | OfficeIMO signature | 1 KiB | 0.51x | 1.73x | 0.10x | 1.73x | 1.02x |
| Verify | OfficeIMO signature | 64 KiB | 0.40x | 0.45x | 0.18x | 0.45x | 0.96x |
| Verify | OfficeIMO signature | 1 MiB | 0.69x | 0.036x | 1.03x | 0.30x | 0.97x |
| Verify | Platform signature | 1 KiB | 1.05x | **2.09x** | 0.10x | **2.07x** | 1.02x |
| Verify | Platform signature | 64 KiB | 0.93x | 0.40x | 0.18x | 0.40x | 0.97x |
| Verify | Platform signature | 1 MiB | 0.81x | 0.030x | 1.03x | 0.25x | 0.96x |

Absolute median process peaks are 39.6-44.3 MiB across all lanes. Retained
managed growth is no more than 1.95 KiB for OfficeIMO and no more than 16.34
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
