namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
    /// <summary>Analyzes an existing PDF byte array for profile-specific readback evidence.</summary>
    public static PdfComplianceReadinessReport AssessReadback(PdfComplianceProfile profile, byte[] pdf, PdfReadOptions? options = null) {
        Guard.ComplianceProfile(profile, nameof(profile));
        Guard.NotNull(pdf, nameof(pdf));
        PdfDocumentProbe probe = PdfInspector.Probe(pdf);
        PdfReadDocument document = PdfReadDocument.Open(pdf, options);
        PdfDocumentInfo info = PdfInspector.FromReadDocument(document, probe);
        return AssessReadback(profile, info, document.ExtractAttachments());
    }

    /// <summary>Analyzes an existing PDF file for profile-specific readback evidence.</summary>
    public static PdfComplianceReadinessReport AssessReadback(PdfComplianceProfile profile, string path, PdfReadOptions? options = null) {
        Guard.NotNullOrWhiteSpace(path, nameof(path));
        return AssessReadback(profile, File.ReadAllBytes(path), options);
    }

    /// <summary>Analyzes an existing PDF stream from its current position for profile-specific readback evidence.</summary>
    public static PdfComplianceReadinessReport AssessReadback(PdfComplianceProfile profile, Stream stream, PdfReadOptions? options = null) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(stream));
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return AssessReadback(profile, buffer.ToArray(), options);
    }

    /// <summary>Analyzes already-inspected PDF metadata for profile-specific readback evidence.</summary>
    public static PdfComplianceReadinessReport AssessReadback(PdfComplianceProfile profile, PdfDocumentInfo info) {
        return AssessReadback(profile, info, extractedAttachments: null);
    }

    internal static PdfComplianceReadinessReport AssessReadback(
        PdfComplianceProfile profile,
        PdfReadDocument document,
        PdfDocumentInfo info) {
        Guard.NotNull(document, nameof(document));
        return AssessReadback(profile, info, document.ExtractAttachments());
    }

    private static PdfComplianceReadinessReport AssessReadback(PdfComplianceProfile profile, PdfDocumentInfo info, IReadOnlyList<PdfExtractedAttachment>? extractedAttachments) {
        Guard.ComplianceProfile(profile, nameof(profile));
        Guard.NotNull(info, nameof(info));

        var requirements = new List<PdfComplianceRequirement>();
        if (profile == PdfComplianceProfile.None) {
            return new PdfComplianceReadinessReport(profile, GetDisplayName(profile), requirements.AsReadOnly());
        }

        if (RequiresPdf17FileVersion(profile)) {
            Add(requirements, "readback-pdf-file-version", "Readback effective PDF 1.7 version",
                ComparePdfVersion(info.EffectiveVersion, "1.7") == 0,
                "The saved PDF effective version is PDF 1.7.",
                "Generate the saved PDF with a PDF 1.7 header or catalog /Version before checking PDF/A-2, PDF/A-3, PDF/UA-1, or e-invoice profile evidence.");
        }

        if (RequiresPdf20FileVersion(profile)) {
            Add(requirements, "readback-pdf-file-version", "Readback effective PDF 2.0 version",
                ComparePdfVersion(info.EffectiveVersion, "2.0") >= 0,
                "The saved PDF effective version is PDF 2.0 or newer.",
                "Generate the saved PDF with a PDF 2.0 header or catalog /Version before checking PDF/A-4 or PDF/UA-2 profile evidence.");
        }

        if (IsPdfA(profile) || IsElectronicInvoice(profile)) {
            AddPdfAReadbackRequirements(requirements, profile, info);
        }

        if (RequiresAccessibility(profile)) {
            AddAccessibilityReadbackRequirements(requirements, profile, info);
        }

        if (IsElectronicInvoice(profile)) {
            AddElectronicInvoiceReadbackRequirements(requirements, info, extractedAttachments);
        }

        return new PdfComplianceReadinessReport(profile, GetDisplayName(profile), requirements.AsReadOnly());
    }

    private static void AddPdfAReadbackRequirements(List<PdfComplianceRequirement> requirements, PdfComplianceProfile profile, PdfDocumentInfo info) {
        PdfXmpMetadataInfo? xmp = info.XmpMetadata;
        (int Part, string? Conformance) target = GetPdfAIdentificationTarget(profile);

        Add(requirements, "readback-xmp-metadata", "Readback catalog XMP metadata",
            xmp != null && xmp.IsWellFormedXml,
            "The saved PDF contains readable, well-formed catalog XMP metadata.",
            "The saved PDF must contain readable, well-formed catalog XMP metadata.");

        bool hasMatchingIdentification = xmp != null &&
            xmp.PdfAPart == target.Part &&
            IsPdfAConformanceMatch(xmp.PdfAConformance, target.Conformance);
        Add(requirements, "readback-pdfa-identification", "Readback PDF/A identification XMP",
            hasMatchingIdentification,
            "The saved PDF contains PDF/A identification metadata for " + GetDisplayName(profile) + ".",
            "The saved PDF XMP metadata must contain matching pdfaid:part and pdfaid:conformance values for " + GetDisplayName(profile) + ".");

        Add(requirements, "readback-output-intent", "Readback catalog output intent",
            info.OutputIntentCount > 0,
            "The saved PDF contains at least one readable catalog output intent.",
            "The saved PDF must contain a readable catalog output intent.");

        requirements.Add(BuildReadbackOutputIntentPolicyRequirement(info));
        requirements.Add(new PdfComplianceRequirement(
            "verapdf-validation",
            "veraPDF validation evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "Run veraPDF against the saved PDF before claiming PDF/A or PDF/A-backed e-invoice conformance."));
    }

    private static PdfComplianceRequirement BuildReadbackOutputIntentPolicyRequirement(PdfDocumentInfo info) {
        PdfOutputIntentInfo? intent = info.OutputIntents.FirstOrDefault(outputIntent =>
            string.Equals(outputIntent.OutputConditionIdentifier, PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, StringComparison.Ordinal));
        if (intent == null) {
            return new PdfComplianceRequirement(
                "readback-output-intent-policy",
                "Readback sRGB output-intent policy",
                PdfComplianceRequirementStatus.Missing,
                "The saved PDF does not contain an sRGB IEC61966-2.1 output intent.");
        }

        if (!string.Equals(intent.Subtype, "GTS_PDFA1", StringComparison.Ordinal)) {
            return new PdfComplianceRequirement(
                "readback-output-intent-policy",
                "Readback sRGB output-intent policy",
                PdfComplianceRequirementStatus.Missing,
                "The saved PDF sRGB output intent must use /S /GTS_PDFA1 for PDF/A workflows.");
        }

        if (intent.DestinationOutputProfileColorComponents != 3 ||
            intent.DestinationOutputProfileHasIccSignature != true) {
            return new PdfComplianceRequirement(
                "readback-output-intent-policy",
                "Readback sRGB output-intent policy",
                PdfComplianceRequirementStatus.Missing,
                "The saved PDF sRGB output intent must reference an RGB ICC profile with a readable ICC signature.");
        }

        return new PdfComplianceRequirement(
            "readback-output-intent-policy",
            "Readback sRGB output-intent policy",
            PdfComplianceRequirementStatus.Satisfied,
            "The saved PDF contains an sRGB IEC61966-2.1 /GTS_PDFA1 output intent with RGB ICC profile evidence.");
    }

    private static int ComparePdfVersion(string? left, string? right) {
        if (!TryParsePdfVersion(left, out int leftMajor, out int leftMinor)) {
            return TryParsePdfVersion(right, out _, out _) ? -1 : 0;
        }

        if (!TryParsePdfVersion(right, out int rightMajor, out int rightMinor)) {
            return 1;
        }

        int majorComparison = leftMajor.CompareTo(rightMajor);
        return majorComparison != 0 ? majorComparison : leftMinor.CompareTo(rightMinor);
    }

    private static bool TryParsePdfVersion(string? version, out int major, out int minor) {
        major = 0;
        minor = 0;
        if (string.IsNullOrWhiteSpace(version)) {
            return false;
        }

        string[] parts = version!.Split('.');
        return parts.Length == 2 &&
            int.TryParse(parts[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out major) &&
            int.TryParse(parts[1], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out minor);
    }

    private static void AddAccessibilityReadbackRequirements(List<PdfComplianceRequirement> requirements, PdfComplianceProfile profile, PdfDocumentInfo info) {
        PdfXmpMetadataInfo? xmp = info.XmpMetadata;
        if (profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2) {
            int expectedPart = profile == PdfComplianceProfile.PdfUa2 ? 2 : 1;
            Add(requirements, "readback-pdfua-identification", "Readback PDF/UA identification XMP",
                xmp?.PdfUaPart == expectedPart,
                "The saved PDF contains " + GetDisplayName(profile) + " identification metadata.",
                "The saved PDF XMP metadata must contain pdfuaid:part=" + expectedPart.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");

            Add(requirements, "readback-document-title", "Readback document title metadata",
                !string.IsNullOrWhiteSpace(xmp?.Title),
                "The saved PDF contains a non-empty XMP dc:title.",
                "The saved PDF XMP metadata must contain a non-empty dc:title.");

            Add(requirements, "readback-display-document-title", "Readback viewer displays document title",
                info.ViewerPreferences?.GetBoolean("DisplayDocTitle") == true,
                "The saved PDF catalog ViewerPreferences dictionary sets DisplayDocTitle true.",
                "The saved PDF catalog ViewerPreferences dictionary must set DisplayDocTitle true.");

            requirements.Add(new PdfComplianceRequirement(
                "pdfua-validation",
                "PDF/UA validator evidence",
                PdfComplianceRequirementStatus.Unsupported,
                "Run a PDF/UA validator against the saved PDF before claiming " + GetDisplayName(profile) + " conformance."));
        }

        Add(requirements, "readback-document-language", "Readback document language",
            IsValidPdfLanguageTag(info.CatalogLanguage),
            "The saved PDF catalog contains a valid document language.",
            "The saved PDF catalog must contain a valid /Lang entry such as en-US.");

        PdfTaggedContentInfo? tagged = info.TaggedContent;
        Add(requirements, "readback-marked-catalog", "Readback marked catalog",
            tagged?.Marked == true,
            "The saved PDF catalog /MarkInfo dictionary sets /Marked true.",
            "The saved PDF catalog must contain /MarkInfo with /Marked true.");

        Add(requirements, "readback-structure-root", "Readback structure tree root",
            tagged?.StructTreeRootObjectNumber.HasValue == true,
            "The saved PDF catalog references a readable StructTreeRoot.",
            "The saved PDF catalog must reference a readable StructTreeRoot.");

        Add(requirements, "readback-parent-tree-next-key", "Readback parent-tree next key",
            tagged?.ParentTreeNextKey.HasValue == true,
            "The saved PDF StructTreeRoot contains ParentTreeNextKey.",
            "The saved PDF StructTreeRoot must contain ParentTreeNextKey for generated parent-tree entries.");

        bool hasDocumentElement = tagged != null &&
            tagged.StructureElements.Any(static element => string.Equals(element.StructureType, "Document", StringComparison.Ordinal));
        Add(requirements, "readback-document-structure-element", "Readback document structure element",
            hasDocumentElement,
            "The saved PDF contains a readable /Document structure element.",
            "The saved PDF structure tree must contain a readable /Document structure element.");

        Add(requirements, "readback-structure-element-count", "Readback structure element count",
            tagged != null && tagged.StructureElementCount > 0,
            "The saved PDF contains " + (tagged?.StructureElementCount ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture) + " readable structure element(s).",
            "The saved PDF structure tree must contain at least one readable structure element.");

        Add(requirements, "readback-marked-content-references", "Readback marked-content references",
            tagged != null && tagged.MarkedContentReferenceCount > 0,
            "The saved PDF contains " + (tagged?.MarkedContentReferenceCount ?? 0).ToString(System.Globalization.CultureInfo.InvariantCulture) + " structure-tree marked-content reference(s).",
            "The saved PDF structure tree must contain marked-content references that connect structure elements to page content.");

        bool hasReadableStructureEvidence = tagged?.Marked == true &&
            tagged.StructTreeRootObjectNumber.HasValue &&
            tagged.ParentTreeNextKey.HasValue &&
            hasDocumentElement &&
            tagged.StructureElementCount > 0 &&
            tagged.MarkedContentReferenceCount > 0;
        Add(requirements, "readback-tagged-structure-evidence", "Readback tagged-structure evidence",
            hasReadableStructureEvidence,
            "The saved PDF exposes a marked catalog, structure root, parent tree, document element, structure elements, and marked-content references for exact-artifact validation.",
            "The saved PDF must expose a marked catalog, structure root, parent tree, document element, structure elements, and marked-content references before external PDF/UA validation.");
    }

    private static void AddElectronicInvoiceReadbackRequirements(List<PdfComplianceRequirement> requirements, PdfDocumentInfo info, IReadOnlyList<PdfExtractedAttachment>? extractedAttachments) {
        PdfXmpMetadataInfo? xmp = info.XmpMetadata;
        Add(requirements, "readback-einvoice-xmp", "Readback e-invoice XMP metadata",
            xmp?.HasElectronicInvoiceMetadata == true,
            "The saved PDF contains Factur-X/ZUGFeRD XMP extension metadata.",
            "The saved PDF XMP metadata must contain Factur-X/ZUGFeRD extension properties.");

        requirements.Add(BuildReadbackInvoiceAttachmentRequirement(extractedAttachments));

        requirements.Add(new PdfComplianceRequirement(
            "mustang-validation",
            "Mustang validation evidence",
            PdfComplianceRequirementStatus.Unsupported,
            "Run Mustang against the saved PDF before claiming Factur-X or ZUGFeRD conformance."));
    }

    private static PdfComplianceRequirement BuildReadbackInvoiceAttachmentRequirement(IReadOnlyList<PdfExtractedAttachment>? attachments) {
        if (attachments == null) {
            return new PdfComplianceRequirement(
                "readback-associated-invoice-file",
                "Readback associated invoice file",
                PdfComplianceRequirementStatus.Missing,
                "Analyze PDF bytes, a file path, or a stream so OfficeIMO.Pdf can validate the factur-x.xml CrossIndustryInvoice payload during readback.");
        }

        var diagnostics = new List<string>();
        for (int i = 0; i < attachments.Count; i++) {
            PdfExtractedAttachment attachment = attachments[i];
            if (!TryCreateReadbackEmbeddedFile(attachment, diagnostics, out PdfEmbeddedFile? embeddedFile)) {
                continue;
            }

            if (IsFacturXCiiAttachment(embeddedFile!, diagnostics)) {
                return new PdfComplianceRequirement(
                    "readback-associated-invoice-file",
                    "Readback associated invoice file",
                    PdfComplianceRequirementStatus.Satisfied,
                    "The saved PDF contains canonical factur-x.xml associated-file evidence with parseable UN/CEFACT CrossIndustryInvoice XML.");
            }
        }

        string diagnostic = diagnostics.Count == 0
            ? "The saved PDF must contain factur-x.xml with an associated-file relationship, XML MIME type, and parseable UN/CEFACT CrossIndustryInvoice XML."
            : string.Join(" ", diagnostics.Distinct(StringComparer.Ordinal).ToArray());
        return new PdfComplianceRequirement(
            "readback-associated-invoice-file",
            "Readback associated invoice file",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static bool TryCreateReadbackEmbeddedFile(PdfExtractedAttachment attachment, List<string> diagnostics, out PdfEmbeddedFile? embeddedFile) {
        embeddedFile = null;
        byte[] bytes = attachment.Bytes;
        if (bytes.Length == 0) {
            diagnostics.Add("Attach non-empty UN/CEFACT CrossIndustryInvoice XML in factur-x.xml.");
            return false;
        }

        try {
            embeddedFile = new PdfEmbeddedFile(
                attachment.UnicodeFileName ?? attachment.FileName,
                bytes,
                attachment.MimeType,
                attachment.Relationship,
                attachment.Description);
            return true;
        } catch (ArgumentException ex) {
            diagnostics.Add(ex.Message);
            return false;
        }
    }
}
