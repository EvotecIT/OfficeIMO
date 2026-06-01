namespace OfficeIMO.Pdf;

internal static class PdfComplianceValidator {
    private static readonly string[] PdfABaseRequirements = {
        "profile-required output-intent validation and approved ICC profile policy",
        "profile-specific XMP identification metadata",
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

        if (options.ComplianceProfile == PdfComplianceProfile.None) {
            return;
        }

        throw new NotSupportedException(BuildUnsupportedProfileMessage(options.ComplianceProfile));
    }

    private static string BuildUnsupportedProfileMessage(PdfComplianceProfile profile) {
        return "PDF compliance profile " + GetDisplayName(profile) + " was requested, but OfficeIMO.Pdf cannot yet generate certified " + GetProfileFamily(profile) + " output. Missing generated-profile support: " + string.Join("; ", GetRequirements(profile)) + ". Use " + nameof(PdfComplianceProfile) + "." + nameof(PdfComplianceProfile.None) + " until the profile is implemented, or use the existing XMP metadata, ToUnicode, and embedded-font options only as compliance groundwork without claiming formal conformance.";
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

        if (RequiresElectronicInvoice(profile)) {
            foreach (string requirement in ElectronicInvoiceRequirements) {
                yield return requirement;
            }
        }
    }

    private static IEnumerable<string> GetBaseRequirements(PdfComplianceProfile profile) =>
        RequiresElectronicInvoice(profile)
            ? PdfABaseRequirements.Concat(UnicodeRequirements)
            : profile == PdfComplianceProfile.PdfUa1
                ? System.Array.Empty<string>()
                : PdfABaseRequirements;

    private static bool RequiresUnicodeMapping(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2U ||
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3U ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfUa1;

    private static bool RequiresAccessibility(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfUa1;

    private static bool RequiresElectronicInvoice(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.FacturX ||
        profile == PdfComplianceProfile.Zugferd;

    private static string GetProfileFamily(PdfComplianceProfile profile) {
        if (profile == PdfComplianceProfile.PdfUa1) {
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
            PdfComplianceProfile.PdfUa1 => "PDF/UA-1",
            PdfComplianceProfile.FacturX => "Factur-X",
            PdfComplianceProfile.Zugferd => "ZUGFeRD",
            _ => "None"
        };
}
