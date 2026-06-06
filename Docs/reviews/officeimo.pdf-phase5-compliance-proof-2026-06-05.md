# OfficeIMO.Pdf Phase 5 Compliance Proof Notes

Date: 2026-06-05
Branch: `codex/pdf-implementation-review-20260605`

## Scope

Phase 5 separates OfficeIMO.Pdf readiness from formal conformance claims.
OfficeIMO.Pdf can inspect its own options and generated evidence, but PDF/A,
PDF/UA, Factur-X, and ZUGFeRD still need independent validator evidence before
the library or a caller should claim conformance.

## Added proof model

- `PdfExternalValidationResult` records caller- or CI-supplied validator output.
- `PdfExternalValidatorKind` models veraPDF, PDF/UA validators, Mustang, and custom validator families.
- `PdfExternalValidationStatus` distinguishes not-run, passed, failed, and validator-error results.
- `PdfComplianceAnalyzer.AssessProof(...)` combines the existing readiness report with external validator evidence.
- `PdfComplianceProofReport.CanClaimConformance` is true only when internal readiness passes and every required external validator family has a passing result without a failed/error result.

## Required validator families

| Profile family | Required external evidence |
| --- | --- |
| PDF/A-2 and PDF/A-3 | veraPDF |
| PDF/UA-1 | PDF/UA validator |
| Factur-X and ZUGFeRD | veraPDF and Mustang |

`PdfComplianceAnalyzer.Assess(...)` remains a readiness API. The new proof API is
the safer gate for callers that want to display, export, or enforce conformance
claims.

## Existing validator workflow

The test suite already has optional external validator gates:

- `OFFICEIMO_VERAPDF` or `OFFICEIMO_VERAPDF_PATH`
- `OFFICEIMO_VERAPDF_ARGS` to adapt local veraPDF command-line syntax
- `OFFICEIMO_PDFUA_VALIDATOR` or `OFFICEIMO_PDFUA_VALIDATOR_PATH`
- `OFFICEIMO_PDFUA_VALIDATOR_ARGS` to adapt the chosen PDF/UA validator command-line syntax
- `OFFICEIMO_MUSTANG` or `OFFICEIMO_MUSTANG_PATH`
- `OFFICEIMO_MUSTANG_ARGS` to adapt local Mustang command-line syntax
- `OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS=1` to fail when validators are unavailable

Those tests remain responsible for process execution. The product layer accepts
validator results rather than launching tools itself, keeping OfficeIMO.Pdf usable
in desktop apps, services, CI, and conversion pipelines with different validator
installation models.

## Artifact workflow

`Build/Export-PdfComplianceProof.ps1` now turns the optional gate tests into a
repeatable proof pack:

```powershell
Build/Export-PdfComplianceProof.ps1 -OutputDirectory artifacts/pdf-compliance-proof -Framework net8.0
```

The script sets `OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT`, runs
`PdfComplianceGateTests`, and writes:

- `officeimo-pdfa3-groundwork.pdf`
- `officeimo-pdfua-groundwork.pdf`
- `officeimo-einvoice-groundwork.pdf`
- validator diagnostics for veraPDF, the configured PDF/UA validator, and
  Mustang, including not-run diagnostics when validators are unavailable
- `officeimo-profile-proof-contract.json` with the
  `PdfComplianceAnalyzer.AssessProof(...)` engine-level claimability view
- `index.md` for human review
- `proof.json` with schema version `3`, commit, command, test exit code, contract
  flags, validator configuration booleans, artifact hashes, validator
  kind/profile/status metadata, expected validator status, expected-status
  match flags, validator diagnostic hashes, profile-level proof matrix rows,
  product proof contract rows, and working-tree state

Use `-RequireValidators` when CI should fail if veraPDF, the configured PDF/UA
validator, or Mustang are missing. Without that switch, missing validators are
recorded as proof-pack diagnostics. The script also accepts `-VeraPdfPath`,
`-VeraPdfArgs`, `-PdfUaValidatorPath`, `-PdfUaValidatorArgs`, `-MustangPath`,
and `-MustangArgs` as direct wrappers over the existing validator environment
variables for release runners.

`.github/workflows/pdf-compliance-proof.yml` runs this proof pack on PDF
compliance changes, validates the pack with
`.github/scripts/Assert-PdfComplianceProof.ps1`, uploads the proof artifact, and
adds a GitHub step summary with fixture hashes, observed validator statuses,
expected statuses, match flags, the profile proof matrix, and the product proof
contract. It also offers manual `require_validators`, `verapdf_path`, `verapdf_args`,
`pdfua_validator_path`, `pdfua_validator_args`, `mustang_path`, and
`mustang_args` dispatch inputs for strict validator environments.

For the current groundwork fixtures, `NotRun` is expected when a validator is
not configured, and `Failed` is expected when a validator is configured and
runs. A `Passed` validator result is intentionally unexpected until formal
profile generation is implemented and the relevant gate is flipped.

The profile matrix gives release tooling stable rows for `pdfa-3b-groundwork`,
`pdfua-1-groundwork`, and `einvoice-groundwork`, including fixture file,
validator diagnostic file, readiness requirement id, observed/expected status,
and `canClaimConformance=false`.
The product proof contract records the same fail-closed state from
`PdfComplianceAnalyzer.AssessProof(...)`, including internal readiness, missing
external validator families, unsupported requirement ids, and
`canClaimConformance=false` for every current groundwork profile. Its
`externalEvidenceMode=NoExternalValidationInjected` makes the evidence boundary
explicit: validator diagnostics are interpreted separately by the proof script.

## Next useful work

- Configure veraPDF, the chosen PDF/UA validator, and Mustang on the intended strict CI runner and run the proof workflow with `require_validators`.
- Replace the generic PDF/UA validator command with the team's selected validator package/install step once CI ownership is decided.
- Wire conversion examples to call `AssessProof(...)` before displaying formal compliance badges.
- Keep expanding internal readiness until fewer requirements are placeholders, then make validator failures easier to map back to layout/content diagnostics.
