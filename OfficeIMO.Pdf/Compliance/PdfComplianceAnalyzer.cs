namespace OfficeIMO.Pdf;

/// <summary>
/// Analyzes generated-PDF options against the requirements of planned formal compliance profiles.
/// </summary>
internal static partial class PdfComplianceAnalyzer {
    /// <summary>Analyzes the compliance profile requested by the supplied options.</summary>
    public static PdfComplianceReadinessReport Assess(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        return Assess(options.ComplianceProfile, options);
    }

    /// <summary>Analyzes the supplied options against a requested formal compliance profile.</summary>
    public static PdfComplianceReadinessReport Assess(PdfComplianceProfile profile, PdfOptions options) {
        return Assess(profile, options, generatedStandardFonts: null);
    }

    /// <summary>
    /// Analyzes the supplied options against a requested formal compliance profile, including generated standard-font usage evidence when available.
    /// </summary>
    public static PdfComplianceReadinessReport Assess(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfStandardFont>? generatedStandardFonts) {
        return AssessCore(profile, options, generatedStandardFonts, generatedFontUsages: null, documentTitle: null, hasDocumentMetadataEvidence: false, generatedImages: null, generatedDrawings: null, generatedForms: null);
    }

    internal static PdfComplianceReadinessReport AssessDocument(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfStandardFont>? generatedStandardFonts, IEnumerable<PdfGeneratedFontComplianceEvidence>? generatedFontUsages, string? documentTitle, IEnumerable<PdfGeneratedImageAccessibilityEvidence>? generatedImages, IEnumerable<PdfGeneratedDrawingAccessibilityEvidence>? generatedDrawings, IEnumerable<PdfGeneratedFormAccessibilityEvidence>? generatedForms) {
        return AssessCore(profile, options, generatedStandardFonts, generatedFontUsages, documentTitle, hasDocumentMetadataEvidence: true, generatedImages: generatedImages, generatedDrawings: generatedDrawings, generatedForms: generatedForms);
    }

    private static PdfComplianceReadinessReport AssessCore(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfStandardFont>? generatedStandardFonts, IEnumerable<PdfGeneratedFontComplianceEvidence>? generatedFontUsages, string? documentTitle, bool hasDocumentMetadataEvidence, IEnumerable<PdfGeneratedImageAccessibilityEvidence>? generatedImages, IEnumerable<PdfGeneratedDrawingAccessibilityEvidence>? generatedDrawings, IEnumerable<PdfGeneratedFormAccessibilityEvidence>? generatedForms) {
        Guard.ComplianceProfile(profile, nameof(profile));
        Guard.NotNull(options, nameof(options));

        PdfStandardFont[]? generatedFontSnapshot = SnapshotGeneratedStandardFonts(generatedStandardFonts);
        PdfGeneratedFontComplianceEvidence[]? generatedFontUsageSnapshot = SnapshotGeneratedFontUsages(generatedFontUsages);
        PdfGeneratedImageAccessibilityEvidence[]? generatedImageSnapshot = SnapshotGeneratedImages(generatedImages);
        PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawingSnapshot = SnapshotGeneratedDrawings(generatedDrawings);
        PdfGeneratedFormAccessibilityEvidence[]? generatedFormSnapshot = SnapshotGeneratedForms(generatedForms);
        var requirements = new List<PdfComplianceRequirement>();
        if (profile == PdfComplianceProfile.None) {
            return new PdfComplianceReadinessReport(profile, GetDisplayName(profile), requirements.AsReadOnly());
        }

        if (RequiresPdf17FileVersion(profile)) {
            AddFileVersionRequirement(requirements, options);
        }

        if (RequiresPdf20FileVersion(profile)) {
            AddPdf20FileVersionRequirement(requirements, options);
        }

        if (IsPdfA(profile) || IsElectronicInvoice(profile)) {
            AddPdfARequirements(requirements, profile, options, generatedFontSnapshot, generatedFontUsageSnapshot);
        }

        if (RequiresUnicodeMapping(profile) || IsElectronicInvoice(profile)) {
            AddUnicodeRequirements(requirements, options, generatedFontUsageSnapshot);
        }

        if (RequiresAccessibility(profile)) {
            AddAccessibilityRequirements(requirements, profile, options, documentTitle, hasDocumentMetadataEvidence, generatedImageSnapshot, generatedDrawingSnapshot, generatedFormSnapshot);
        }

        if (IsElectronicInvoice(profile)) {
            AddElectronicInvoiceRequirements(requirements, options);
        }

        return new PdfComplianceReadinessReport(profile, GetDisplayName(profile), requirements.AsReadOnly());
    }

    private static void AddFileVersionRequirement(List<PdfComplianceRequirement> requirements, PdfOptions options) {
        Add(requirements, "pdf-file-version", "PDF 1.7 file header",
            options.FileVersion == PdfFileVersion.Pdf17,
            "Generated output is configured for a PDF 1.7 file header.",
            "Set PdfOptions.FileVersion or PdfDocument.FileVersion(...) to PdfFileVersion.Pdf17 for PDF/A-2, PDF/A-3, PDF/UA-1, and e-invoice profile groundwork.");
    }

    private static void AddUnicodeRequirements(List<PdfComplianceRequirement> requirements, PdfOptions options, PdfGeneratedFontComplianceEvidence[]? generatedFontUsages) {
        bool hasEmbeddedUnicodeCoverage = HasEmbeddedGeneratedFontUnicodeCoverage(generatedFontUsages);
        Add(requirements, "standard-font-to-unicode", "Standard-font ToUnicode maps",
            options.IncludeStandardFontToUnicodeMaps || hasEmbeddedUnicodeCoverage,
            "Generated standard-font resources will include WinAnsi ToUnicode CMaps.",
            hasEmbeddedUnicodeCoverage
                ? "Generated text uses embedded TrueType Type0 fonts with Identity-H ToUnicode CMaps."
                : "Enable PdfOptions.IncludeStandardFontToUnicodeMaps for non-embedded Type1 standard-font resources, or embed every generated font slot.");

        requirements.Add(BuildFullUnicodeMappingRequirement(generatedFontUsages, hasEmbeddedUnicodeCoverage));
    }

    private static PdfComplianceRequirement BuildFullUnicodeMappingRequirement(PdfGeneratedFontComplianceEvidence[]? generatedFontUsages, bool hasEmbeddedUnicodeCoverage) {
        if (generatedFontUsages == null) {
            return new PdfComplianceRequirement(
                "full-unicode-mapping",
                "Full generated text Unicode mapping",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated text usage evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to verify whether every generated text font slot uses embedded Type0 Unicode mapping.");
        }

        if (hasEmbeddedUnicodeCoverage) {
            return new PdfComplianceRequirement(
                "full-unicode-mapping",
                "Full generated text Unicode mapping",
                PdfComplianceRequirementStatus.Satisfied,
                generatedFontUsages.Length == 0
                    ? "No generated text font usage was reported for this document."
                    : "Every generated text font slot reported by layout uses a parseable embedded TrueType or OpenType/CFF Type0 font with Identity-H ToUnicode mapping.");
        }

        return new PdfComplianceRequirement(
            "full-unicode-mapping",
            "Full generated text Unicode mapping",
            PdfComplianceRequirementStatus.Missing,
            "Embed parseable TrueType or OpenType/CFF mappings for every generated text font slot before claiming full generated text Unicode mapping.");
    }

    private static bool HasEmbeddedGeneratedFontUnicodeCoverage(PdfGeneratedFontComplianceEvidence[]? generatedFontUsages) {
        if (generatedFontUsages == null) {
            return false;
        }

        if (generatedFontUsages.Length == 0) {
            return true;
        }

        for (int i = 0; i < generatedFontUsages.Length; i++) {
            PdfGeneratedFontComplianceEvidence usage = generatedFontUsages[i];
            IReadOnlyDictionary<PdfStandardFont, PdfEmbeddedFont> embeddedFonts = usage.Options.EmbeddedFonts;
            if (!embeddedFonts.TryGetValue(usage.Font, out PdfEmbeddedFont? embeddedFont)) {
                return false;
            }

            if (!TryParseEmbeddedFont(embeddedFont, out _)) {
                return false;
            }
        }

        return true;
    }

    private static PdfStandardFont[]? SnapshotGeneratedStandardFonts(IEnumerable<PdfStandardFont>? generatedStandardFonts) {
        if (generatedStandardFonts == null) {
            return null;
        }

        var fonts = new HashSet<PdfStandardFont>();
        foreach (PdfStandardFont font in generatedStandardFonts) {
            Guard.StandardFont(font, nameof(generatedStandardFonts), "Generated standard-font usage contains an unsupported PDF font.");
            fonts.Add(font);
        }

        return fonts
            .OrderBy(font => (int)font)
            .ToArray();
    }

    private static PdfGeneratedFontComplianceEvidence[]? SnapshotGeneratedFontUsages(IEnumerable<PdfGeneratedFontComplianceEvidence>? generatedFontUsages) {
        if (generatedFontUsages == null) {
            return null;
        }

        return generatedFontUsages.ToArray();
    }

    private static PdfGeneratedImageAccessibilityEvidence[]? SnapshotGeneratedImages(IEnumerable<PdfGeneratedImageAccessibilityEvidence>? generatedImages) {
        if (generatedImages == null) {
            return null;
        }

        return generatedImages.ToArray();
    }

    private static PdfGeneratedDrawingAccessibilityEvidence[]? SnapshotGeneratedDrawings(IEnumerable<PdfGeneratedDrawingAccessibilityEvidence>? generatedDrawings) {
        if (generatedDrawings == null) {
            return null;
        }

        return generatedDrawings.ToArray();
    }

    private static PdfGeneratedFormAccessibilityEvidence[]? SnapshotGeneratedForms(IEnumerable<PdfGeneratedFormAccessibilityEvidence>? generatedForms) {
        if (generatedForms == null) {
            return null;
        }

        return generatedForms.ToArray();
    }

    private static void Add(List<PdfComplianceRequirement> requirements, string id, string displayName, bool satisfied, string satisfiedDiagnostic, string missingDiagnostic) {
        requirements.Add(new PdfComplianceRequirement(
            id,
            displayName,
            satisfied ? PdfComplianceRequirementStatus.Satisfied : PdfComplianceRequirementStatus.Missing,
            satisfied ? satisfiedDiagnostic : missingDiagnostic));
    }

    private static bool IsPdfA(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2B ||
        profile == PdfComplianceProfile.PdfA2U ||
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3B ||
        profile == PdfComplianceProfile.PdfA3U ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfA4 ||
        profile == PdfComplianceProfile.PdfA4E ||
        profile == PdfComplianceProfile.PdfA4F;

    private static bool RequiresPdf17FileVersion(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA2B ||
        profile == PdfComplianceProfile.PdfA2U ||
        profile == PdfComplianceProfile.PdfA2A ||
        profile == PdfComplianceProfile.PdfA3B ||
        profile == PdfComplianceProfile.PdfA3U ||
        profile == PdfComplianceProfile.PdfA3A ||
        profile == PdfComplianceProfile.PdfUa1 ||
        IsElectronicInvoice(profile);

    private static bool RequiresPdf20FileVersion(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.PdfA4 ||
        profile == PdfComplianceProfile.PdfA4E ||
        profile == PdfComplianceProfile.PdfA4F ||
        profile == PdfComplianceProfile.PdfUa2;

    private static void AddPdf20FileVersionRequirement(List<PdfComplianceRequirement> requirements, PdfOptions options) {
        Add(requirements, "pdf-file-version", "PDF 2.0 file header",
            options.FileVersion == PdfFileVersion.Pdf20,
            "Generated output is configured for a PDF 2.0 file header.",
            "Set PdfOptions.FileVersion or PdfDocument.FileVersion(...) to PdfFileVersion.Pdf20 for PDF/A-4 and PDF/UA-2 groundwork.");
    }

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

    private static bool IsElectronicInvoice(PdfComplianceProfile profile) =>
        profile == PdfComplianceProfile.FacturX ||
        profile == PdfComplianceProfile.Zugferd;

    private static string GetDisplayName(PdfComplianceProfile profile) {
        switch (profile) {
            case PdfComplianceProfile.PdfA2B:
                return "PDF/A-2b";
            case PdfComplianceProfile.PdfA2U:
                return "PDF/A-2u";
            case PdfComplianceProfile.PdfA2A:
                return "PDF/A-2a";
            case PdfComplianceProfile.PdfA3B:
                return "PDF/A-3b";
            case PdfComplianceProfile.PdfA3U:
                return "PDF/A-3u";
            case PdfComplianceProfile.PdfA3A:
                return "PDF/A-3a";
            case PdfComplianceProfile.PdfA4:
                return "PDF/A-4";
            case PdfComplianceProfile.PdfA4E:
                return "PDF/A-4e";
            case PdfComplianceProfile.PdfA4F:
                return "PDF/A-4f";
            case PdfComplianceProfile.PdfUa1:
                return "PDF/UA-1";
            case PdfComplianceProfile.PdfUa2:
                return "PDF/UA-2";
            case PdfComplianceProfile.FacturX:
                return "Factur-X";
            case PdfComplianceProfile.Zugferd:
                return "ZUGFeRD";
            default:
                return "None";
        }
    }
}
