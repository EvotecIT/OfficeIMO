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

    private static readonly string[] PdfXRequirements = {
        "profile-specific pdfxid XMP identification",
        "CMYK print-condition ICC output intent with /GTS_PDFX subtype",
        "embedded-font coverage for every generated glyph",
        "black-preserving source-color conversion and exact color-space proof",
        "profile-specific transparency, annotation, form, layer, page-box, and self-containment policy",
        "qualified external PDF/X preflight validation bound to the exact artifact"
    };

    internal static void ValidateGenerationOptions(PdfCreateOptions options) {
        Guard.NotNull(options, nameof(options));
        Guard.ComplianceProfile(options.ComplianceProfile, nameof(options.ComplianceProfile));

        foreach (PdfComplianceProfile profile in RequestedProfiles(options)) {
            if (!IsSupported(profile)) throw new NotSupportedException(BuildUnsupportedProfileMessage(profile));
        }
    }

    // All declared A/UA claims participate in the gate, including combined output.
    private static IEnumerable<PdfComplianceProfile> RequestedProfiles(PdfCreateOptions options) {
        if (options.ComplianceProfile == PdfComplianceProfile.None) yield break;
        var profiles = new HashSet<PdfComplianceProfile> { options.ComplianceProfile };
        if (options.PdfAIdentification is PdfAIdentification a) {
            profiles.Add(a.Part switch {
                2 => a.Conformance == "A" ? PdfComplianceProfile.PdfA2A : a.Conformance == "U" ? PdfComplianceProfile.PdfA2U : PdfComplianceProfile.PdfA2B,
                3 => a.Conformance == "A" ? PdfComplianceProfile.PdfA3A : a.Conformance == "U" ? PdfComplianceProfile.PdfA3U : PdfComplianceProfile.PdfA3B,
                _ => a.Conformance == "E" ? PdfComplianceProfile.PdfA4E : a.Conformance == "F" ? PdfComplianceProfile.PdfA4F : PdfComplianceProfile.PdfA4
            });
        }
        if (options.PdfUaIdentification is PdfUaIdentification ua)
            profiles.Add(ua.Part == 2 ? PdfComplianceProfile.PdfUa2 : PdfComplianceProfile.PdfUa1);
        if (options.ElectronicInvoiceMetadata != null) profiles.Add(PdfComplianceProfile.FacturX);
        foreach (PdfComplianceProfile profile in profiles) yield return profile;
    }

    private static bool IsSupported(PdfComplianceProfile profile) =>
        profile is PdfComplianceProfile.PdfA2B or PdfComplianceProfile.PdfA2U or PdfComplianceProfile.PdfA2A
            or PdfComplianceProfile.PdfA3B or PdfComplianceProfile.PdfA3U or PdfComplianceProfile.PdfA3A
            or PdfComplianceProfile.PdfA4 or PdfComplianceProfile.PdfA4E or PdfComplianceProfile.PdfA4F
            or PdfComplianceProfile.PdfUa1 or PdfComplianceProfile.PdfUa2
            or PdfComplianceProfile.FacturX or PdfComplianceProfile.Zugferd
            or PdfComplianceProfile.PdfX1A2003 or PdfComplianceProfile.PdfX4;

    internal static void ValidateGeneratedDocument(PdfCreateOptions options, string? documentTitle, PdfGeneratedDocumentComplianceEvidence evidence) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(evidence, nameof(evidence));
        foreach (PdfComplianceProfile profile in RequestedProfiles(options)) {
            ThrowForGaps(profile, PdfComplianceAnalyzer.AssessDocument(profile, options, evidence, documentTitle), "generation");
        }
    }

    internal static bool RequiresExactArtifactValidation(PdfCreateOptions options) => options.ComplianceProfile != PdfComplianceProfile.None;

    /// <summary>Prevents native save options from retaining identification that prohibits encryption.</summary>
    internal static void ValidateOutputEncryptionIdentification(PdfReadDocument document) {
        PdfXmpMetadataInfo? xmp = document.XmpMetadata;
        PdfMetadata metadata = document.Metadata;
        if (xmp?.HasPdfAIdentification == true || xmp?.HasPdfXIdentification == true ||
            xmp?.HasElectronicInvoiceMetadata == true || !string.IsNullOrEmpty(metadata.PdfXVersion) ||
            !string.IsNullOrEmpty(metadata.PdfXConformance))
            throw new ArgumentException("PDF Standard encryption cannot retain PDF/A, PDF/X, Factur-X, or ZUGFeRD identification.");
    }

    internal static void ValidateGeneratedArtifact(
        PdfCreateOptions options, byte[] artifact,
        PdfLoadOptions? readOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(artifact, nameof(artifact));
        cancellationToken.ThrowIfCancellationRequested();
        foreach (PdfComplianceProfile profile in RequestedProfiles(options)) {
            ThrowForGaps(profile, PdfComplianceAnalyzer.AssessReadback(profile, artifact, cancellationToken, readOptions), "exact-artifact");
        }
    }

    private static void ThrowForGaps(PdfComplianceProfile profile, PdfComplianceReadinessReport readiness, string stage) {
        PdfComplianceRequirement[] gaps = readiness.Requirements.Where(requirement =>
            !PdfComplianceProofReport.IsExternalValidationRequirement(requirement.Id) &&
            requirement.Status != PdfComplianceRequirementStatus.Satisfied).ToArray();
        if (gaps.Length == 0) return;
        throw new InvalidOperationException(GetDisplayName(profile) + " " + stage + " requirements are not satisfied: " +
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

        if (profile == PdfComplianceProfile.PdfX1A2003 || profile == PdfComplianceProfile.PdfX4) {
            foreach (string requirement in PdfXRequirements) {
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

    private static bool IsPdfX(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfX1A2003 ||
        profile == PdfComplianceProfile.PdfX4;

    private static string GetProfileFamily(PdfComplianceProfile profile) {
        if (profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2) {
            return "PDF/UA";
        }

        if (RequiresElectronicInvoice(profile)) {
            return "EN 16931 e-invoice";
        }

        if (profile == PdfComplianceProfile.PdfX1A2003 || profile == PdfComplianceProfile.PdfX4) {
            return "PDF/X";
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
            PdfComplianceProfile.PdfX1A2003 => "PDF/X-1a:2003",
            PdfComplianceProfile.PdfX4 => "PDF/X-4",
            _ => "None"
        };
}
