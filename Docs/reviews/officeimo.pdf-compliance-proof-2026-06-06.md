# OfficeIMO.Pdf Compliance Proof Workflow

Date: 2026-06-06
Branch: `codex/pdf-premium-roadmap-20260606`

## Purpose

This proof lane makes compliance status reviewable without turning on formal
claims too early. The generated PDFs are groundwork fixtures only. They prove
that current PDF/A-3b, PDF/UA-1, and e-invoice primitives can be emitted and
externally checked, while `ComplianceProfile != None` remains blocked until a
profile has internal readiness plus passing validator evidence.

## Command

```powershell
Build/Export-PdfComplianceProof.ps1 -OutputDirectory artifacts/pdf-compliance-proof -Framework net8.0
```

Use `-RequireValidators` when a CI or release lane must fail if veraPDF, the
configured PDF/UA validator, or Mustang are unavailable.

Strict or release runs can also pass validator locations directly:

```powershell
Build/Export-PdfComplianceProof.ps1 -OutputDirectory artifacts/pdf-compliance-proof -Framework net8.0 -RequireValidators -VeraPdfPath /opt/verapdf/verapdf -PdfUaValidatorPath /opt/pdfua/validator -MustangPath /opt/mustang/mustangproject.jar
```

## Outputs

- `officeimo-pdfa3-groundwork.pdf`
- `officeimo-pdfua-groundwork.pdf`
- `officeimo-einvoice-groundwork.pdf`
- `verapdf-pdfa3-groundwork.txt`
- `pdfua-groundwork.txt`
- `mustang-einvoice-groundwork.txt`
- `officeimo-profile-proof-contract.json`
- `index.md`
- `proof.json`

The JSON summary records schema version `3`, commit, command, test exit code,
strict-validator mode, validator configuration booleans, contract flags, PDF
fixture hashes, validator kind, profile, observed status, expected status,
expected-status match, diagnostic hashes, profile-level proof matrix rows, and
the `PdfComplianceAnalyzer.AssessProof(...)` product proof contract with
internal readiness, missing validators, and `canClaimConformance=false` rows for
the current groundwork profiles, plus `externalEvidenceMode` and working tree
state. The product proof contract intentionally uses
`NoExternalValidationInjected`; external validator execution evidence is tracked
separately through the validator diagnostic rows.

## CI Lane

`.github/workflows/pdf-compliance-proof.yml` runs the proof script on PDF
compliance changes, validates the proof pack with
`.github/scripts/Assert-PdfComplianceProof.ps1`, and uploads the generated PDFs,
diagnostics, `officeimo-profile-proof-contract.json`, `index.md`, and
`proof.json` as the
`officeimo-pdf-compliance-proof` artifact.

The workflow also writes a GitHub step summary with fixture hashes, observed
validator statuses, expected statuses, match flags, profile proof rows, and the
engine-derived product proof contract, so reviewers can see whether veraPDF, the
PDF/UA validator, and Mustang were not run, failed as expected, or unexpectedly
passed without downloading the artifact first.

The default PR/push lane records missing validator diagnostics without failing.
The manual `workflow_dispatch` input `require_validators` runs the same script
with `-RequireValidators` for CI environments where veraPDF, the PDF/UA
validator, and Mustang are expected to be installed. Manual dispatch also
accepts `verapdf_path`, `verapdf_args`, `pdfua_validator_path`,
`pdfua_validator_args`, `mustang_path`, and `mustang_args`, which are passed
through to the proof script for strict runners that install validator tools
outside `PATH` or need command-line syntax overrides.

## Validator Configuration

- `OFFICEIMO_VERAPDF` or `OFFICEIMO_VERAPDF_PATH`
- `OFFICEIMO_VERAPDF_ARGS`
- `OFFICEIMO_PDFUA_VALIDATOR` or `OFFICEIMO_PDFUA_VALIDATOR_PATH`
- `OFFICEIMO_PDFUA_VALIDATOR_ARGS`
- `OFFICEIMO_MUSTANG` or `OFFICEIMO_MUSTANG_PATH`
- `OFFICEIMO_MUSTANG_ARGS`
- `OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS=1`

The script sets `OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT` for the test process.
Normal test runs stay artifact-free.

`Build/Export-PdfComplianceProof.ps1` also accepts `-VeraPdfPath`,
`-VeraPdfArgs`, `-PdfUaValidatorPath`, `-PdfUaValidatorArgs`, `-MustangPath`,
and `-MustangArgs`; these are thin wrappers over the same environment variables
and are restored after the test process exits.

## Current Local Proof

The first local smoke run used:

```powershell
Build/Export-PdfComplianceProof.ps1 -OutputDirectory artifacts/pdf-compliance-proof-test -Framework net8.0 -NoRestore
```

Result: `PdfComplianceGateTests` passed with `11/11` tests. The artifact pack
contained three PDFs, three validator diagnostics,
`officeimo-profile-proof-contract.json`, `index.md`, and `proof.json`.
veraPDF, the PDF/UA validator, and Mustang were not configured in this
environment, so the diagnostics recorded not-run status rather than external
validator output. The proof JSON also recorded that no validator executable or
custom args were configured. Each diagnostic recorded `expectedStatus=NotRun`
and `matchesExpectedStatus=true`; when a validator is configured, the current
groundwork fixtures expect `Failed` until formal conformance generation is
implemented. The `profileProofs` matrix maps each profile family to its
groundwork fixture, validator diagnostic, readiness requirement id, expected
status, and `canClaimConformance=false`.
The `productProofContract` section records the same fail-closed claim state from
`PdfComplianceAnalyzer.AssessProof(...)`: PDF/A-3b is missing veraPDF, PDF/UA-1
is missing the configured PDF/UA validator, and Factur-X is missing veraPDF and
Mustang. Its `externalEvidenceMode=NoExternalValidationInjected` makes clear
that this is the engine readiness contract before validator diagnostics are
interpreted by the proof script.

## Next Gap

The next premium compliance step is to install or configure veraPDF, the chosen
PDF/UA validator, and Mustang in the intended strict CI lane, run the same proof
script with `-RequireValidators`, and map the external failures back into the
existing readiness requirement IDs.
