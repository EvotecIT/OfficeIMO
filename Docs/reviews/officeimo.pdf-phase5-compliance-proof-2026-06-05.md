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
- `OFFICEIMO_MUSTANG` or `OFFICEIMO_MUSTANG_PATH`
- `OFFICEIMO_REQUIRE_PDF_COMPLIANCE_VALIDATORS=1` to fail when validators are unavailable

Those tests remain responsible for process execution. The product layer accepts
validator results rather than launching tools itself, keeping OfficeIMO.Pdf usable
in desktop apps, services, CI, and conversion pipelines with different validator
installation models.

## Next useful work

- Emit a machine-readable proof artifact, such as JSON, from the compliance gate tests.
- Add a PDF/UA external validator lane once the team chooses the validator used in CI.
- Wire conversion examples to call `AssessProof(...)` before displaying formal compliance badges.
- Keep expanding internal readiness until fewer requirements are placeholders, then make validator failures easier to map back to layout/content diagnostics.
