namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {

    private static void AddAccessibilityRequirements(List<PdfComplianceRequirement> requirements, PdfComplianceProfile profile, PdfOptions options, string? documentTitle, bool hasDocumentMetadataEvidence, PdfGeneratedImageAccessibilityEvidence[]? generatedImages, PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings, PdfGeneratedFormAccessibilityEvidence[]? generatedForms) {
        if (profile == PdfComplianceProfile.PdfUa1 || profile == PdfComplianceProfile.PdfUa2) {
            AddPdfUaIdentificationRequirement(requirements, profile, options);
            requirements.Add(BuildDocumentTitleRequirement(options, documentTitle, hasDocumentMetadataEvidence));
            requirements.Add(BuildDisplayDocumentTitleRequirement(options));
            requirements.Add(new PdfComplianceRequirement(
                "pdfua-validation",
                "PDF/UA validator evidence",
                PdfComplianceRequirementStatus.Unsupported,
                "External PDF/UA validator evidence has not been supplied. Use PdfComplianceAnalyzer.AssessProof(...) with a passing PDF/UA validator result before claiming " + GetDisplayName(profile) + "."));
        }

        requirements.Add(BuildDocumentLanguageRequirement(options));

        Add(requirements, "tagged-catalog-markers", "Tagged PDF catalog markers",
            options.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers,
            "Catalog /MarkInfo and /StructTreeRoot markers are configured as tagged-PDF groundwork.",
            "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() to emit tagged-PDF catalog markers. This is groundwork only; full structure-tree content generation is still required.");
        requirements.Add(BuildTaggedPageTabOrderRequirement(options));
        requirements.Add(BuildTaggedParentTreeNextKeyRequirement(options));
        requirements.Add(BuildGeneratedDocumentStructureRootRequirement(options));
        requirements.Add(BuildGeneratedDocumentStructureLanguageRequirement(options));

        requirements.Add(BuildTaggedStructureRequirement(options, hasDocumentMetadataEvidence, generatedImages, generatedDrawings, generatedForms));
        requirements.Add(BuildGeneratedTextBlockStructureReferenceRequirement(options));
        requirements.Add(BuildGeneratedListStructureReferenceRequirement(options));
        requirements.Add(BuildGeneratedListContainerStructureRequirement(options));
        requirements.Add(BuildGeneratedTableCellStructureReferenceRequirement(options));
        requirements.Add(BuildGeneratedTableContainerStructureRequirement(options));
        requirements.Add(BuildGeneratedTableHeaderScopeRequirement(options));
        requirements.Add(BuildGeneratedTableSpanAttributeRequirement(options));
        requirements.Add(BuildGeneratedTableCaptionStructureRequirement(options));
        requirements.Add(BuildGeneratedLinkAnnotationStructureRequirement(options));
        requirements.Add(BuildGeneratedLinkTextStructureRequirement(options));
        requirements.Add(BuildGeneratedFormWidgetStructureRequirement(options, generatedForms));
        requirements.Add(BuildGeneratedFormFieldAccessibleNameRequirement(generatedForms));
        requirements.Add(BuildGeneratedImageAlternativeTextRequirement(generatedImages));
        requirements.Add(BuildGeneratedImageStructureReferenceRequirement(options, generatedImages));
        requirements.Add(BuildGeneratedDecorativeImageArtifactRequirement(generatedImages));
        requirements.Add(BuildGeneratedDrawingAlternativeTextRequirement(generatedDrawings));
        requirements.Add(BuildGeneratedDrawingStructureReferenceRequirement(options, generatedDrawings));
        requirements.Add(BuildGeneratedDecorativeDrawingArtifactRequirement(options, generatedDrawings));
        requirements.Add(BuildGeneratedRunningPageTextArtifactRequirement(options));
        requirements.Add(BuildGeneratedDecorativeFlowRuleArtifactRequirement(options));
        requirements.Add(BuildGeneratedDecorativeLayoutArtifactRequirement(options));
        requirements.Add(BuildAlternateTextRequirement(generatedImages, generatedDrawings));
    }

    private static PdfComplianceRequirement BuildTaggedPageTabOrderRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "tagged-page-tab-order",
                "Tagged page tab order",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated page dictionaries can emit /Tabs /S and follow structure-tree order for tab navigation.");
        }

        return new PdfComplianceRequirement(
            "tagged-page-tab-order",
            "Tagged page tab order",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged generated pages will emit /Tabs /S so annotation and keyboard tab order follows the structure tree as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildTaggedStructureRequirement(PdfOptions options, bool hasDocumentMetadataEvidence, PdfGeneratedImageAccessibilityEvidence[]? generatedImages, PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings, PdfGeneratedFormAccessibilityEvidence[]? generatedForms) {
        if (!hasDocumentMetadataEvidence || generatedImages == null || generatedDrawings == null || generatedForms == null) {
            return new PdfComplianceRequirement(
                "tagged-structure",
                "Tagged PDF structure tree",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated structure evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to verify OfficeIMO-generated structure coverage; external PDF/UA/PDF/A-a validation is still required before claiming conformance.");
        }

        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "tagged-structure",
                "Tagged PDF structure tree",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated content can emit a tagged structure tree with parent-tree references.");
        }

        if (!IsValidPdfLanguageTag(options.Language)) {
            return new PdfComplianceRequirement(
                "tagged-structure",
                "Tagged PDF structure tree",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.Language or PdfDocument.Language(...) to a valid language tag so generated /Document structure elements can carry /Lang metadata.");
        }

        int missingImageAlternativeText = generatedImages.Count(image => !image.IsDecorativeArtifact && !image.HasAlternativeText);
        int missingDrawingAlternativeText = generatedDrawings.Count(drawing => !drawing.IsDecorativeArtifact && !drawing.HasAlternativeText);
        int missingFormAccessibleNames = generatedForms.Count(form => !form.HasAccessibleName);
        if (missingImageAlternativeText > 0 || missingDrawingAlternativeText > 0 || missingFormAccessibleNames > 0) {
            return new PdfComplianceRequirement(
                "tagged-structure",
                "Tagged PDF structure tree",
                PdfComplianceRequirementStatus.Missing,
                "Generated tagged structure is enabled, but meaningful generated images, drawings, or form fields still need alternate text/accessibility names before the generated structure coverage can be treated as complete groundwork.");
        }

        return new PdfComplianceRequirement(
            "tagged-structure",
            "Tagged PDF structure tree",
            PdfComplianceRequirementStatus.Satisfied,
            "OfficeIMO-generated content has tagged-structure groundwork: /Document structure root, document language metadata, parent-tree references, paragraph and heading marked content, list containers, table containers and cell references, link text and annotation references, form widget references, image and drawing figure/artifact handling, decorative page text/artifact markers, and structure-order tab hints. External PDF/UA/PDF/A-a validation is still required before claiming formal conformance.");
    }

    private static PdfComplianceRequirement BuildTaggedParentTreeNextKeyRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "tagged-parent-tree-next-key",
                "Tagged parent-tree next key",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated StructTreeRoot dictionaries can emit /ParentTreeNextKey for generated parent-tree entries.");
        }

        return new PdfComplianceRequirement(
            "tagged-parent-tree-next-key",
            "Tagged parent-tree next key",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit /ParentTreeNextKey on generated StructTreeRoot dictionaries as PDF/UA/PDF/A-a groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedDocumentStructureRootRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-document-structure-root",
                "Generated document structure root",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated top-level structure elements can be nested below a /Document structure element.");
        }

        return new PdfComplianceRequirement(
            "generated-document-structure-root",
            "Generated document structure root",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will nest generated top-level structure elements below a generated /Document structure element as PDF/UA/PDF/A-a groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedDocumentStructureLanguageRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-document-structure-language",
                "Generated document structure language",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated /Document structure elements can carry document language metadata.");
        }

        if (string.IsNullOrWhiteSpace(options.Language)) {
            return new PdfComplianceRequirement(
                "generated-document-structure-language",
                "Generated document structure language",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.Language or PdfDocument.Language(...) so generated /Document structure elements can emit /Lang metadata.");
        }

        if (!IsValidPdfLanguageTag(options.Language)) {
            return new PdfComplianceRequirement(
                "generated-document-structure-language",
                "Generated document structure language",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.Language or PdfDocument.Language(...) to a valid language tag such as en-US before emitting generated /Document /Lang metadata.");
        }

        return new PdfComplianceRequirement(
            "generated-document-structure-language",
            "Generated document structure language",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit /Lang metadata on generated /Document structure elements as PDF/UA/PDF/A-a groundwork.");
    }

    private static PdfComplianceRequirement BuildDocumentLanguageRequirement(PdfOptions options) {
        if (string.IsNullOrWhiteSpace(options.Language)) {
            return new PdfComplianceRequirement(
                "document-language",
                "Document language",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.Language or PdfDocument.Language(...) for accessibility-oriented profiles.");
        }

        if (!IsValidPdfLanguageTag(options.Language)) {
            return new PdfComplianceRequirement(
                "document-language",
                "Document language",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.Language or PdfDocument.Language(...) to a valid language tag such as en-US before claiming accessibility-oriented profiles.");
        }

        return new PdfComplianceRequirement(
            "document-language",
            "Document language",
            PdfComplianceRequirementStatus.Satisfied,
            "A catalog document language is configured with valid language-tag syntax.");
    }

    private static bool IsValidPdfLanguageTag(string? language) {
        if (string.IsNullOrWhiteSpace(language)) {
            return false;
        }

        string value = language!.Trim();
        string[] parts = value.Split('-');
        if (parts.Length == 0 || !IsAsciiLetterSubtag(parts[0], 2, 3)) {
            return false;
        }

        for (int i = 1; i < parts.Length; i++) {
            if (!IsAsciiAlphaNumericSubtag(parts[i], 1, 8)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsAsciiLetterSubtag(string value, int minLength, int maxLength) {
        if (value.Length < minLength || value.Length > maxLength) {
            return false;
        }

        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            if (!((character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z'))) {
                return false;
            }
        }

        return true;
    }

    private static bool IsAsciiAlphaNumericSubtag(string value, int minLength, int maxLength) {
        if (value.Length < minLength || value.Length > maxLength) {
            return false;
        }

        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            if (!((character >= 'A' && character <= 'Z') ||
                  (character >= 'a' && character <= 'z') ||
                  (character >= '0' && character <= '9'))) {
                return false;
            }
        }

        return true;
    }

    private static void AddPdfUaIdentificationRequirement(List<PdfComplianceRequirement> requirements, PdfComplianceProfile profile, PdfOptions options) {
        PdfUaIdentification? identification = options.PdfUaIdentification;
        int expectedPart = profile == PdfComplianceProfile.PdfUa2 ? 2 : 1;
        Add(requirements, "pdfua-identification", "PDF/UA identification XMP",
            identification != null && identification.Part == expectedPart,
            "PDF/UA identification metadata matches " + GetDisplayName(profile) + ".",
            "Set PdfOptions.SetPdfUaIdentification(" + expectedPart.ToString(System.Globalization.CultureInfo.InvariantCulture) + ") before claiming " + GetDisplayName(profile) + ".");
    }

    private static PdfComplianceRequirement BuildDocumentTitleRequirement(PdfOptions options, string? documentTitle, bool hasDocumentMetadataEvidence) {
        if (!hasDocumentMetadataEvidence) {
            return new PdfComplianceRequirement(
                "document-title",
                "Document title metadata",
                PdfComplianceRequirementStatus.Unsupported,
                "Document title metadata was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include document metadata evidence.");
        }

        bool emitsXmpMetadata = WillEmitXmpMetadata(options);
        if (!string.IsNullOrWhiteSpace(documentTitle) && emitsXmpMetadata) {
            return new PdfComplianceRequirement(
                "document-title",
                "Document title metadata",
                PdfComplianceRequirementStatus.Satisfied,
                "A non-empty document title will be emitted in XMP dc:title metadata.");
        }

        string diagnostic = string.IsNullOrWhiteSpace(documentTitle)
            ? "Set a non-empty PdfDocument.Meta(title: ...) value before claiming accessibility-oriented profiles."
            : "Enable XMP metadata emission with PdfOptions.IncludeXmpMetadata or profile identification metadata so the document title is emitted as dc:title.";
        return new PdfComplianceRequirement(
            "document-title",
            "Document title metadata",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildDisplayDocumentTitleRequirement(PdfOptions options) {
        PdfViewerPreferencesOptions? viewerPreferences = options.ViewerPreferences;
        if (viewerPreferences != null && viewerPreferences.DisplayDocTitle == true) {
            return new PdfComplianceRequirement(
                "display-document-title",
                "Viewer displays document title",
                PdfComplianceRequirementStatus.Satisfied,
                "Catalog viewer preferences will set DisplayDocTitle to true.");
        }

        return new PdfComplianceRequirement(
            "display-document-title",
            "Viewer displays document title",
            PdfComplianceRequirementStatus.Missing,
            "Set PdfOptions.ViewerPreferences.DisplayDocTitle or configure PdfDocument.ViewerPreferences(...) so the catalog ViewerPreferences dictionary includes DisplayDocTitle true.");
    }

    private static PdfComplianceRequirement BuildGeneratedImageAlternativeTextRequirement(PdfGeneratedImageAccessibilityEvidence[]? generatedImages) {
        if (generatedImages == null) {
            return new PdfComplianceRequirement(
                "generated-image-alternate-text",
                "Generated image alternate text",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated image evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated image evidence.");
        }

        PdfGeneratedImageAccessibilityEvidence[] meaningfulImages = generatedImages
            .Where(image => !image.IsDecorativeArtifact)
            .ToArray();

        if (meaningfulImages.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-image-alternate-text",
                "Generated image alternate text",
                PdfComplianceRequirementStatus.Satisfied,
                "No non-decorative generated images were reported for this document.");
        }

        int missing = meaningfulImages.Count(image => !image.HasAlternativeText);
        if (missing == 0) {
            return new PdfComplianceRequirement(
                "generated-image-alternate-text",
                "Generated image alternate text",
                PdfComplianceRequirementStatus.Satisfied,
                "Every non-decorative generated image reported by layout has alternate text.");
        }

        string diagnostic = missing == meaningfulImages.Length
            ? "Set PdfImageStyle.AlternativeText or pass alternativeText to Image(...) for generated meaningful images, including header/footer images."
            : "Set PdfImageStyle.AlternativeText or pass alternativeText to Image(...) for the " + missing.ToString(System.Globalization.CultureInfo.InvariantCulture) + " non-decorative generated image(s) missing alternate text.";
        return new PdfComplianceRequirement(
            "generated-image-alternate-text",
            "Generated image alternate text",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildGeneratedDecorativeImageArtifactRequirement(PdfGeneratedImageAccessibilityEvidence[]? generatedImages) {
        if (generatedImages == null) {
            return new PdfComplianceRequirement(
                "decorative-image-artifacts",
                "Decorative image artifact markers",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated image evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated decorative-image evidence.");
        }

        int decorativeCount = generatedImages.Count(image => image.IsDecorativeArtifact);
        if (decorativeCount == 0) {
            return new PdfComplianceRequirement(
                "decorative-image-artifacts",
                "Decorative image artifact markers",
                PdfComplianceRequirementStatus.Satisfied,
                "No generated decorative page background or image watermark images were reported for this document.");
        }

        return new PdfComplianceRequirement(
            "decorative-image-artifacts",
            "Decorative image artifact markers",
            PdfComplianceRequirementStatus.Satisfied,
            "Generated decorative page background and image watermark images are emitted as PDF artifact marked content.");
    }

    private static PdfComplianceRequirement BuildGeneratedDrawingAlternativeTextRequirement(PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings) {
        if (generatedDrawings == null) {
            return new PdfComplianceRequirement(
                "generated-drawing-alternate-text",
                "Generated drawing alternate text",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated drawing evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated shape and drawing evidence.");
        }

        PdfGeneratedDrawingAccessibilityEvidence[] meaningfulDrawings = generatedDrawings
            .Where(drawing => !drawing.IsDecorativeArtifact)
            .ToArray();

        if (meaningfulDrawings.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-drawing-alternate-text",
                "Generated drawing alternate text",
                PdfComplianceRequirementStatus.Satisfied,
                "No non-decorative generated shapes or drawing scenes were reported for this document.");
        }

        int missing = meaningfulDrawings.Count(drawing => !drawing.HasAlternativeText);
        if (missing == 0) {
            return new PdfComplianceRequirement(
                "generated-drawing-alternate-text",
                "Generated drawing alternate text",
                PdfComplianceRequirementStatus.Satisfied,
                "Every non-decorative generated shape and drawing scene reported by layout has alternate text.");
        }

        string diagnostic = missing == meaningfulDrawings.Length
            ? "Set PdfDrawingStyle.AlternativeText for generated meaningful shapes and drawing scenes, or set PdfDrawingStyle.Decorative for purely decorative drawing content."
            : "Set PdfDrawingStyle.AlternativeText for the " + missing.ToString(System.Globalization.CultureInfo.InvariantCulture) + " non-decorative generated shape or drawing scene(s) missing alternate text.";
        return new PdfComplianceRequirement(
            "generated-drawing-alternate-text",
            "Generated drawing alternate text",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildGeneratedDrawingStructureReferenceRequirement(PdfOptions options, PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings) {
        if (generatedDrawings == null) {
            return new PdfComplianceRequirement(
                "generated-drawing-structure-references",
                "Generated drawing structure references",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated drawing evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated drawing structure-reference evidence.");
        }

        PdfGeneratedDrawingAccessibilityEvidence[] meaningfulDrawings = generatedDrawings
            .Where(drawing => !drawing.IsDecorativeArtifact)
            .ToArray();
        if (meaningfulDrawings.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-drawing-structure-references",
                "Generated drawing structure references",
                PdfComplianceRequirementStatus.Satisfied,
                "No non-decorative generated shapes or drawing scenes were reported for this document.");
        }

        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-drawing-structure-references",
                "Generated drawing structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated shape and drawing figures can receive MCID and parent-tree structure references.");
        }

        int missingAltText = meaningfulDrawings.Count(drawing => !drawing.HasAlternativeText);
        if (missingAltText > 0) {
            return new PdfComplianceRequirement(
                "generated-drawing-structure-references",
                "Generated drawing structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set alternate text on generated meaningful shapes and drawing scenes before relying on generated drawing figure MCID and parent-tree references.");
        }

        return new PdfComplianceRequirement(
            "generated-drawing-structure-references",
            "Generated drawing structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit MCID-marked generated shape and drawing figures plus StructTreeRoot and ParentTree references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedDecorativeDrawingArtifactRequirement(PdfOptions options, PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings) {
        if (generatedDrawings == null) {
            return new PdfComplianceRequirement(
                "decorative-drawing-artifacts",
                "Decorative drawing artifact markers",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated drawing evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated decorative-drawing evidence.");
        }

        int decorativeCount = generatedDrawings.Count(drawing => drawing.IsDecorativeArtifact);
        if (decorativeCount == 0) {
            return new PdfComplianceRequirement(
                "decorative-drawing-artifacts",
                "Decorative drawing artifact markers",
                PdfComplianceRequirementStatus.Satisfied,
                "No generated decorative shape or drawing scene flow blocks were reported for this document.");
        }

        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "decorative-drawing-artifacts",
                "Decorative drawing artifact markers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated decorative shape and drawing scene flow blocks can be emitted as artifact marked content.");
        }

        return new PdfComplianceRequirement(
            "decorative-drawing-artifacts",
            "Decorative drawing artifact markers",
            PdfComplianceRequirementStatus.Satisfied,
            "Generated decorative shape and drawing scene flow blocks are emitted as PDF artifact marked content when tagged-PDF markers are enabled.");
    }

    private static PdfComplianceRequirement BuildAlternateTextRequirement(PdfGeneratedImageAccessibilityEvidence[]? generatedImages, PdfGeneratedDrawingAccessibilityEvidence[]? generatedDrawings) {
        if (generatedImages == null || generatedDrawings == null) {
            return new PdfComplianceRequirement(
                "alternate-text",
                "Alternate text for meaningful visuals",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated visual evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include image and drawing alternate-text evidence.");
        }

        int missingImages = generatedImages.Count(image => !image.IsDecorativeArtifact && !image.HasAlternativeText);
        int missingDrawings = generatedDrawings.Count(drawing => !drawing.IsDecorativeArtifact && !drawing.HasAlternativeText);
        if (missingImages == 0 && missingDrawings == 0) {
            return new PdfComplianceRequirement(
                "alternate-text",
                "Alternate text for meaningful visuals",
                PdfComplianceRequirementStatus.Satisfied,
                "Every generated non-decorative image, shape, and drawing scene reported by layout has alternate text; decorative visuals are reported as artifacts.");
        }

        var diagnostics = new List<string>();
        if (missingImages > 0) {
            diagnostics.Add(missingImages.ToString(System.Globalization.CultureInfo.InvariantCulture) + " generated image(s)");
        }

        if (missingDrawings > 0) {
            diagnostics.Add(missingDrawings.ToString(System.Globalization.CultureInfo.InvariantCulture) + " generated shape or drawing scene(s)");
        }

        return new PdfComplianceRequirement(
            "alternate-text",
            "Alternate text for meaningful visuals",
            PdfComplianceRequirementStatus.Missing,
            "Set alternate text for " + string.Join(" and ", diagnostics.ToArray()) + " that are not decorative artifacts.");
    }

    private static PdfComplianceRequirement BuildGeneratedRunningPageTextArtifactRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "decorative-running-page-text-artifacts",
                "Decorative running page text artifact markers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated running header/footer text can be emitted as artifact marked content.");
        }

        return new PdfComplianceRequirement(
            "decorative-running-page-text-artifacts",
            "Decorative running page text artifact markers",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated running header/footer text as PDF artifact marked content while header/footer images with alternate text remain meaningful figures.");
    }

    private static PdfComplianceRequirement BuildGeneratedDecorativeFlowRuleArtifactRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "decorative-flow-rule-artifacts",
                "Decorative flow rule artifact markers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated horizontal rules in document and row/column flow can be emitted as artifact marked content.");
        }

        return new PdfComplianceRequirement(
            "decorative-flow-rule-artifacts",
            "Decorative flow rule artifact markers",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated horizontal rules in document and row/column flow as PDF artifact marked content.");
    }

    private static PdfComplianceRequirement BuildGeneratedDecorativeLayoutArtifactRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "decorative-layout-artifacts",
                "Decorative layout artifact markers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated panel and table layout chrome can be emitted as artifact marked content.");
        }

        return new PdfComplianceRequirement(
            "decorative-layout-artifacts",
            "Decorative layout artifact markers",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated panel backgrounds, panel borders, table fills, table borders, row/column separators, and explicit cell borders as PDF artifact marked content.");
    }

    private static PdfComplianceRequirement BuildGeneratedTextBlockStructureReferenceRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-text-structure-references",
                "Generated paragraph and heading structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated paragraphs and headings can receive MCID and parent-tree structure references.");
        }

        return new PdfComplianceRequirement(
            "generated-text-structure-references",
            "Generated paragraph and heading structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit MCID-marked generated paragraph and heading slices plus StructTreeRoot and ParentTree references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedListStructureReferenceRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-list-structure-references",
                "Generated list item structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated list labels and bodies can receive MCID and parent-tree structure references.");
        }

        return new PdfComplianceRequirement(
            "generated-list-structure-references",
            "Generated list item structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit MCID-marked generated list labels and bodies plus StructTreeRoot and ParentTree references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedListContainerStructureRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-list-structure-containers",
                "Generated list structure containers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated list content can be nested below L, LI, Lbl, and LBody structure elements.");
        }

        return new PdfComplianceRequirement(
            "generated-list-structure-containers",
            "Generated list structure containers",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will nest generated list labels and bodies below generated L and LI structure elements as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedTableCellStructureReferenceRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-table-cell-structure-references",
                "Generated table cell structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated table header and data cell slices can receive MCID and parent-tree structure references.");
        }

        return new PdfComplianceRequirement(
            "generated-table-cell-structure-references",
            "Generated table cell structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit MCID-marked generated TH and TD table cell slices plus StructTreeRoot and ParentTree references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedTableContainerStructureRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-table-structure-containers",
                "Generated table structure containers",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated table cell slices can be nested below Table and TR structure elements.");
        }

        return new PdfComplianceRequirement(
            "generated-table-structure-containers",
            "Generated table structure containers",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will nest generated TH and TD table cell slices below generated Table and TR structure elements as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedTableHeaderScopeRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-table-header-scope-attributes",
                "Generated table header scope attributes",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated table header cells can receive table header scope attributes.");
        }

        return new PdfComplianceRequirement(
            "generated-table-header-scope-attributes",
            "Generated table header scope attributes",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated table header cells with /A /Table /Scope /Column attributes as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedTableSpanAttributeRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-table-span-attributes",
                "Generated table span attributes",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated column-spanned and row-spanned table cells can receive table span attributes.");
        }

        return new PdfComplianceRequirement(
            "generated-table-span-attributes",
            "Generated table span attributes",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated column-spanned and row-spanned table cells with /ColSpan and /RowSpan attributes as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedTableCaptionStructureRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-table-caption-structure-references",
                "Generated table caption structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated table captions can receive MCID and parent-tree structure references below generated Table structure elements.");
        }

        return new PdfComplianceRequirement(
            "generated-table-caption-structure-references",
            "Generated table caption structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated table captions as /Caption content below generated Table structure elements as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedLinkAnnotationStructureRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-link-annotation-structure-references",
                "Generated link annotation structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated link annotations can receive /Link OBJR structure references and annotation StructParent entries.");
        }

        return new PdfComplianceRequirement(
            "generated-link-annotation-structure-references",
            "Generated link annotation structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated link annotations with /StructParent entries plus /Link OBJR structure references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedLinkTextStructureRequirement(PdfOptions options) {
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-link-text-structure-references",
                "Generated link text structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated rich text links can combine visible /Link MCID references with annotation OBJR references.");
        }

        return new PdfComplianceRequirement(
            "generated-link-text-structure-references",
            "Generated link text structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit generated rich text links with visible /Link marked-content references combined with annotation OBJR references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedFormWidgetStructureRequirement(PdfOptions options, PdfGeneratedFormAccessibilityEvidence[]? generatedForms) {
        if (generatedForms == null) {
            return new PdfComplianceRequirement(
                "generated-form-widget-structure-references",
                "Generated form widget structure references",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated form widget evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated AcroForm widget evidence.");
        }

        if (generatedForms.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-form-widget-structure-references",
                "Generated form widget structure references",
                PdfComplianceRequirementStatus.Satisfied,
                "No generated AcroForm widgets were reported for this document.");
        }

        int widgetCount = generatedForms.Sum(form => form.WidgetCount);
        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-form-widget-structure-references",
                "Generated form widget structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated AcroForm widgets can receive /Form OBJR structure references and annotation StructParent entries.");
        }

        return new PdfComplianceRequirement(
            "generated-form-widget-structure-references",
            "Generated form widget structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit " + widgetCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " generated AcroForm widget(s) with /StructParent entries plus /Form OBJR structure references as PDF/UA groundwork.");
    }

    private static PdfComplianceRequirement BuildGeneratedFormFieldAccessibleNameRequirement(PdfGeneratedFormAccessibilityEvidence[]? generatedForms) {
        if (generatedForms == null) {
            return new PdfComplianceRequirement(
                "generated-form-field-accessible-names",
                "Generated form field accessible names",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated form field evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated AcroForm field evidence.");
        }

        if (generatedForms.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-form-field-accessible-names",
                "Generated form field accessible names",
                PdfComplianceRequirementStatus.Satisfied,
                "No generated AcroForm fields were reported for this document.");
        }

        int missing = generatedForms.Count(form => !form.HasAccessibleName);
        if (missing == 0) {
            return new PdfComplianceRequirement(
                "generated-form-field-accessible-names",
                "Generated form field accessible names",
                PdfComplianceRequirementStatus.Satisfied,
                "Every generated AcroForm field reported by layout has alternate field name metadata.");
        }

        string diagnostic = missing == generatedForms.Length
            ? "Set PdfFormFieldStyle.AlternateName for generated AcroForm fields so output can emit /TU alternate field names for accessibility-oriented profiles."
            : "Set PdfFormFieldStyle.AlternateName for the " + missing.ToString(System.Globalization.CultureInfo.InvariantCulture) + " generated AcroForm field(s) missing /TU alternate field name metadata.";
        return new PdfComplianceRequirement(
            "generated-form-field-accessible-names",
            "Generated form field accessible names",
            PdfComplianceRequirementStatus.Missing,
            diagnostic);
    }

    private static PdfComplianceRequirement BuildGeneratedImageStructureReferenceRequirement(PdfOptions options, PdfGeneratedImageAccessibilityEvidence[]? generatedImages) {
        if (generatedImages == null) {
            return new PdfComplianceRequirement(
                "generated-image-structure-references",
                "Generated image structure references",
                PdfComplianceRequirementStatus.Unsupported,
                "Generated image evidence was not supplied for this options-only readiness assessment. Use PdfDocument.AssessCompliance(...) to include generated image structure-reference evidence.");
        }

        PdfGeneratedImageAccessibilityEvidence[] meaningfulImages = generatedImages
            .Where(image => !image.IsDecorativeArtifact)
            .ToArray();
        if (meaningfulImages.Length == 0) {
            return new PdfComplianceRequirement(
                "generated-image-structure-references",
                "Generated image structure references",
                PdfComplianceRequirementStatus.Satisfied,
                "No non-decorative generated images were reported for this document.");
        }

        if (options.TaggedStructureMode != PdfTaggedStructureMode.CatalogMarkers) {
            return new PdfComplianceRequirement(
                "generated-image-structure-references",
                "Generated image structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set PdfOptions.TaggedStructureMode or PdfDocument.TaggedPdfCatalogMarkers() so generated image figures can receive MCID and parent-tree structure references.");
        }

        int missingAltText = meaningfulImages.Count(image => !image.HasAlternativeText);
        if (missingAltText > 0) {
            return new PdfComplianceRequirement(
                "generated-image-structure-references",
                "Generated image structure references",
                PdfComplianceRequirementStatus.Missing,
                "Set alternate text on generated meaningful images before relying on generated image figure MCID and parent-tree references.");
        }

        return new PdfComplianceRequirement(
            "generated-image-structure-references",
            "Generated image structure references",
            PdfComplianceRequirementStatus.Satisfied,
            "Tagged output will emit MCID-marked generated image figures plus StructTreeRoot and ParentTree references as PDF/UA groundwork.");
    }

    private static bool WillEmitXmpMetadata(PdfOptions options) {
        return options.IncludeXmpMetadata ||
            options.PdfAIdentification != null ||
            options.PdfUaIdentification != null ||
            options.ElectronicInvoiceMetadata != null;
    }
}
