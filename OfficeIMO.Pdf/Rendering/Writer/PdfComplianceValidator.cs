namespace OfficeIMO.Pdf;

internal static class PdfComplianceValidator {
    private static readonly string[] PdfABaseRequirements = {
        "profile-required output-intent validation and approved ICC profile policy",
        "automatic profile-specific XMP identification and extension metadata policy",
        "embedded-font coverage for every generated glyph",
        "profile-aware PDF version and catalog validation",
        "veraPDF validation fixtures in the build lane"
    };

    private static readonly string[] UnicodeRequirements = {
        "Unicode text mapping for every generated text run"
    };

    private static readonly string[] AccessibilityRequirements = {
        "profile-required document language validation",
        "tagged PDF structure tree",
        "role map and reading order",
        "alternate text for meaningful images and drawings"
    };

    private static readonly string[] PdfUaRequirements = {
        "profile-specific PDF file header and catalog version policy",
        "PDF/UA identification XMP",
        "document title metadata",
        "catalog ViewerPreferences DisplayDocTitle true"
    };

    private static readonly string[] ElectronicInvoiceRequirements = {
        "PDF/A-3 output",
        "embedded EN 16931 XML invoice payload",
        "associated-file and embedded-file catalog entries",
        "Factur-X/ZUGFeRD relationship metadata",
        "Mustang validation fixtures in the build lane"
    };

    internal static void ValidateGenerationOptions(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        Guard.ComplianceProfile(options.ComplianceProfile, nameof(options.ComplianceProfile));

        if (options.ComplianceProfile == PdfComplianceProfile.None ||
            options.ComplianceProfile == PdfComplianceProfile.PdfA2B ||
            options.ComplianceProfile == PdfComplianceProfile.PdfA3B ||
            options.ComplianceProfile == PdfComplianceProfile.PdfUa1 ||
            options.ComplianceProfile == PdfComplianceProfile.FacturX ||
            options.ComplianceProfile == PdfComplianceProfile.Zugferd) {
            return;
        }

        throw new NotSupportedException(BuildUnsupportedProfileMessage(options.ComplianceProfile));
    }

    internal static void ValidateGeneratedDocument(PdfOptions options, string? documentTitle, PdfGeneratedDocumentComplianceEvidence evidence) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(evidence, nameof(evidence));
        if (options.ComplianceProfile != PdfComplianceProfile.PdfA2B &&
            options.ComplianceProfile != PdfComplianceProfile.PdfA3B &&
            options.ComplianceProfile != PdfComplianceProfile.PdfUa1 &&
            options.ComplianceProfile != PdfComplianceProfile.FacturX &&
            options.ComplianceProfile != PdfComplianceProfile.Zugferd) {
            return;
        }

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.AssessDocument(
            options.ComplianceProfile,
            options,
            evidence.StandardFonts,
            evidence.FontUsages,
            documentTitle,
            evidence.Images,
            evidence.Drawings,
            evidence.Forms);
        PdfComplianceRequirement[] gaps = readiness.Requirements
            .Where(requirement =>
                !PdfComplianceProofReport.IsExternalValidationRequirement(requirement.Id) &&
                requirement.Status != PdfComplianceRequirementStatus.Satisfied)
            .ToArray();
        if (gaps.Length == 0) {
            return;
        }

        throw new InvalidOperationException(
            GetDisplayName(options.ComplianceProfile) + " generation requirements are not satisfied: " +
            string.Join("; ", gaps.Select(static requirement => requirement.Id + ": " + requirement.Diagnostic)));
    }

    private static string BuildUnsupportedProfileMessage(PdfComplianceProfile profile) {
        return "PDF compliance profile " + GetDisplayName(profile) + " was requested, but OfficeIMO.Pdf cannot yet generate certified " + GetProfileFamily(profile) + " output. Missing generated-profile support: " + string.Join("; ", GetRequirements(profile)) + ". Use " + nameof(PdfComplianceProfile) + "." + nameof(PdfComplianceProfile.None) + " until the profile is implemented, or use the existing file-version, XMP metadata, built-in sRGB output intent, ToUnicode, embedded-font, and e-invoice metadata/attachment options only as compliance groundwork without claiming formal conformance.";
    }

    private static IEnumerable<string> GetRequirements(PdfComplianceProfile profile) {
        foreach (string requirement in GetBaseRequirements(profile)) {
            yield return requirement;
        }

        if (RequiresUnicodeMapping(profile)) {
            foreach (string requirement in UnicodeRequirements) {
                yield return requirement;
            }
        }

        if (RequiresAccessibility(profile)) {
            foreach (string requirement in AccessibilityRequirements) {
                yield return requirement;
            }
        }

        if (profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2) {
            foreach (string requirement in PdfUaRequirements) {
                yield return requirement;
            }
        }

        if (RequiresElectronicInvoice(profile)) {
            foreach (string requirement in ElectronicInvoiceRequirements) {
                yield return requirement;
            }
        }
    }

    private static IEnumerable<string> GetBaseRequirements(PdfComplianceProfile profile) =>
        RequiresElectronicInvoice(profile)
            ? PdfABaseRequirements.Concat(UnicodeRequirements)
            : profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2
                ? System.Array.Empty<string>()
                : PdfABaseRequirements;

    private static bool RequiresUnicodeMapping(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2U ||
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3U ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfA4 ||
        profile == PdfComplianceProfile.PdfA4E ||
        profile == PdfComplianceProfile.PdfA4F ||
        profile == PdfComplianceProfile.PdfUa1 ||
        profile == PdfComplianceProfile.PdfUa2;

    private static bool RequiresAccessibility(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfUa1 ||
        profile == PdfComplianceProfile.PdfUa2;

    private static bool RequiresElectronicInvoice(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.FacturX ||
        profile == PdfComplianceProfile.Zugferd;

    private static string GetProfileFamily(PdfComplianceProfile profile) {
        if (profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2) {
            return "PDF/UA";
        }

        if (RequiresElectronicInvoice(profile)) {
            return "EN 16931 e-invoice";
        }

        return "PDF/A";
    }

    private static string GetDisplayName(PdfComplianceProfile profile) =>
        profile switch {
            PdfComplianceProfile.PdfA2B => "PDF/A-2b",
            PdfComplianceProfile.PdfA2U => "PDF/A-2u",
            PdfComplianceProfile.PdfA2A => "PDF/A-2a",
            PdfComplianceProfile.PdfA3B => "PDF/A-3b",
            PdfComplianceProfile.PdfA3U => "PDF/A-3u",
            PdfComplianceProfile.PdfA3A => "PDF/A-3a",
            PdfComplianceProfile.PdfA4 => "PDF/A-4",
            PdfComplianceProfile.PdfA4E => "PDF/A-4e",
            PdfComplianceProfile.PdfA4F => "PDF/A-4f",
            PdfComplianceProfile.PdfUa1 => "PDF/UA-1",
            PdfComplianceProfile.PdfUa2 => "PDF/UA-2",
            PdfComplianceProfile.FacturX => "Factur-X",
            PdfComplianceProfile.Zugferd => "ZUGFeRD",
            _ => "None"
        };
}
