namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    private static void AddPdfXRequirements(
        List<PdfComplianceRequirement> requirements,
        PdfComplianceProfile profile,
        PdfOptions options,
        PdfStandardFont[]? generatedStandardFonts,
        PdfGeneratedFontComplianceEvidence[]? generatedFontUsages,
        PdfGeneratedDocumentComplianceEvidence? generatedEvidence) {
        PdfXIdentification? identification = options.PdfXIdentification;
        string expectedVersion = profile == PdfComplianceProfile.PdfX1A2003
            ? "PDF/X-1a:2003"
            : "PDF/X-4";
        string? expectedConformance = profile == PdfComplianceProfile.PdfX1A2003
            ? "PDF/X-1a:2003"
            : null;

        Add(requirements, "pdfx-xmp-identification", "PDF/X identification XMP",
            identification != null &&
            string.Equals(identification.Version, expectedVersion, StringComparison.Ordinal) &&
            (expectedConformance == null || string.Equals(identification.Conformance, expectedConformance, StringComparison.Ordinal)),
            "PDF/X identification metadata matches " + GetDisplayName(profile) + ".",
            "Set PdfOptions.PdfXIdentification to matching " + GetDisplayName(profile) + " identification metadata.");

        Add(requirements, "pdfx-no-encryption", "PDF/X encryption policy",
            options.EncryptionSnapshot == null,
            "No Standard security encryption is configured.",
            "PDF/X output cannot use Standard security encryption.");

        PdfOutputIntent? outputIntent = options.OutputIntent;
        Add(requirements, "pdfx-output-intent", "PDF/X CMYK output intent",
            outputIntent != null &&
            outputIntent.Subtype == PdfOutputIntentSubtype.GtsPdfX &&
            outputIntent.Policy == PdfOutputIntentPolicy.PdfXPrintCondition &&
            outputIntent.ColorComponents == 4,
            "A /GTS_PDFX output intent with a CMYK print-condition ICC profile is configured.",
            "Configure a /GTS_PDFX output intent with a caller-supplied CMYK print-condition ICC profile.");

        Add(requirements, "pdfx-trapping-status", "PDF/X trapping status",
            options.TrappingStatus.HasValue,
            "A PDF/X trapping status will be written to the Info dictionary.",
            "Set PdfOptions.TrappingStatus to Unknown, False, or True.");

        Add(requirements, "pdfx-vector-color-conversion", "PDF/X vector and text color conversion",
            options.ConvertVectorColorsToPdfXPrintCondition,
            "Generated page-content RGB color operators will be converted through the CMYK output profile.",
            "Enable PdfOptions.ConvertVectorColorsToPdfXPrintCondition to convert generated page-content RGB color operators through the CMYK output profile.");

        Add(requirements, "pdfx-black-preservation", "PDF/X black preservation",
            options.ConvertVectorColorsToPdfXPrintCondition && options.BlackPreservationMode != PdfBlackPreservationMode.None,
            "Generated vector and text colors use the configured black-preservation policy.",
            "Select PureBlack or NeutralAxis black preservation for generated vector and text color conversion.");

        Add(requirements, "pdfx-raster-color-conversion", "PDF/X raster color conversion",
            options.ConvertRasterImagesToPdfXPrintCondition,
            "Supported generated raster images will be converted through the CMYK output profile.",
            "Enable PdfOptions.ConvertRasterImagesToPdfXPrintCondition to convert supported generated raster images through the CMYK output profile.");

        bool hasCompatibleRasterTransparencyPolicy =
            profile != PdfComplianceProfile.PdfX1A2003 || options.FlattenRasterTransparencyForPdfX;
        Add(requirements, "pdfx-raster-transparency", "PDF/X raster transparency policy",
            hasCompatibleRasterTransparencyPolicy,
            profile == PdfComplianceProfile.PdfX1A2003
                ? "Raster alpha will be flattened against the configured PDF/X background."
                : "Raster alpha is allowed for PDF/X-4 and may be preserved or flattened.",
            "PDF/X-1a raster alpha must be flattened against an explicit background.");

        AddEmbeddedFontCoverageRequirement(requirements, options, generatedStandardFonts, generatedFontUsages);

        requirements.Add(new PdfComplianceRequirement(
            "pdfx-source-color-management",
            "PDF/X source color management",
            PdfComplianceRequirementStatus.Unsupported,
            profile == PdfComplianceProfile.PdfX1A2003
                ? "Appearance streams and every remaining color-bearing resource still require CMYK conversion plus proof that no RGB or transparency remains before PDF/X-1a generation can be enabled."
                : "Appearance-stream color management, vector transparency policy, and exact readback proof are required before PDF/X-4 generation can be enabled."));

        AddGeneratedProductionContentRequirements(requirements, options, generatedEvidence);

        requirements.Add(new PdfComplianceRequirement(
            "pdfx-validation",
            "Qualified PDF/X preflight evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "Run a qualified PDF/X preflight validator against the exact saved artifact before claiming conformance."));
    }

    private static void AddGeneratedProductionContentRequirements(
        List<PdfComplianceRequirement> requirements,
        PdfOptions options,
        PdfGeneratedDocumentComplianceEvidence? evidence) {
        Add(requirements, "pdfx-no-embedded-files", "PDF/X embedded-file policy",
            options.EmbeddedFileSnapshots.Count == 0 && options.PortfolioSnapshot == null,
            "No embedded files or portfolio catalog are configured.",
            "Remove embedded files and portfolio configuration from PDF/X output.");
        Add(requirements, "pdfx-no-catalog-actions", "PDF/X catalog action policy",
            options.OpenActionSnapshot == null && string.IsNullOrWhiteSpace(options.CatalogUriBaseSnapshot),
            "No catalog open action or URI base is configured.",
            "Remove catalog open actions and URI bases from PDF/X output.");

        if (evidence == null) {
            requirements.Add(new PdfComplianceRequirement(
                "pdfx-generated-production-content",
                "Generated PDF/X production content",
                PdfComplianceRequirementStatus.Unsupported,
                "Lay out the generated document to verify page boxes, annotations, forms, layers, and external references."));
            return;
        }

        Add(requirements, "pdfx-page-boxes", "PDF/X TrimBox and BleedBox policy",
            evidence.PageCount > 0 && evidence.PagesWithPrintProductionBoxes == evidence.PageCount,
            "Every generated page has an explicit TrimBox and containing BleedBox policy.",
            "Configure PdfOptions.PrintProductionPageBoxes for every generated page.");
        Add(requirements, "pdfx-no-annotations", "PDF/X annotation policy",
            evidence.AnnotationCount == 0,
            "No page annotations or form widgets will be generated.",
            "Remove links, notes, unflattened visual annotations, and form widgets from PDF/X output.");
        Add(requirements, "pdfx-no-optional-content", "PDF/X optional-content policy",
            evidence.OptionalContentLayerCount == 0,
            "No optional-content layers will be generated.",
            "Remove optional-content layers from PDF/X output.");
        Add(requirements, "pdfx-no-external-references", "PDF/X external-reference policy",
            evidence.ExternalReferenceCount == 0 && string.IsNullOrWhiteSpace(options.CatalogUriBaseSnapshot),
            "No external URI references will be generated.",
            "Remove external URI links and the catalog URI base from PDF/X output.");
    }
}
