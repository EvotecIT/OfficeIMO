namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentConversionResult {
    private static void AddXmpMetadataIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        if (!HasRequiredXmpMetadata(options)) {
            return;
        }

        PdfXmpMetadataInfo? xmp = documentInfo.XmpMetadata;
        if (xmp is null) {
            AddMissingXmpIssues(issues, options);
            return;
        }

        AddStringIssue(issues, "Xmp.Title", options.RequiredXmpTitle, xmp.Title);
        AddStringIssue(issues, "Xmp.Creator", options.RequiredXmpCreator, xmp.Creator);
        AddStringIssue(issues, "Xmp.Description", options.RequiredXmpDescription, xmp.Description);
        AddStringIssue(issues, "Xmp.Producer", options.RequiredXmpProducer, xmp.Producer);
        AddStringIssue(issues, "Xmp.Keywords", options.RequiredXmpKeywords, xmp.Keywords);

        for (int i = 0; i < options.RequiredXmpSubjects.Count; i++) {
            string subject = options.RequiredXmpSubjects[i];
            if (!ContainsExact(xmp.Subjects, subject)) {
                issues.Add(new PdfConversionProofIssue("Xmp.Subject", subject, "missing"));
            }
        }

        if (options.RequiredXmpPdfAPart.HasValue && xmp.PdfAPart != options.RequiredXmpPdfAPart.Value) {
            issues.Add(new PdfConversionProofIssue(
                "Xmp.PdfAPart",
                options.RequiredXmpPdfAPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                xmp.PdfAPart.HasValue ? xmp.PdfAPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty));
        }

        AddStringIssue(issues, "Xmp.PdfAConformance", options.RequiredXmpPdfAConformance, xmp.PdfAConformance);

        if (options.RequiredXmpPdfUaPart.HasValue && xmp.PdfUaPart != options.RequiredXmpPdfUaPart.Value) {
            issues.Add(new PdfConversionProofIssue(
                "Xmp.PdfUaPart",
                options.RequiredXmpPdfUaPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                xmp.PdfUaPart.HasValue ? xmp.PdfUaPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty));
        }
    }

    private static void AddMissingXmpIssues(List<PdfConversionProofIssue> issues, PdfConversionProofOptions options) {
        AddStringIssue(issues, "Xmp.Title", options.RequiredXmpTitle, null);
        AddStringIssue(issues, "Xmp.Creator", options.RequiredXmpCreator, null);
        AddStringIssue(issues, "Xmp.Description", options.RequiredXmpDescription, null);
        AddStringIssue(issues, "Xmp.Producer", options.RequiredXmpProducer, null);
        AddStringIssue(issues, "Xmp.Keywords", options.RequiredXmpKeywords, null);

        for (int i = 0; i < options.RequiredXmpSubjects.Count; i++) {
            issues.Add(new PdfConversionProofIssue("Xmp.Subject", options.RequiredXmpSubjects[i], "missing"));
        }

        if (options.RequiredXmpPdfAPart.HasValue) {
            issues.Add(new PdfConversionProofIssue(
                "Xmp.PdfAPart",
                options.RequiredXmpPdfAPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "missing"));
        }

        AddStringIssue(issues, "Xmp.PdfAConformance", options.RequiredXmpPdfAConformance, null);

        if (options.RequiredXmpPdfUaPart.HasValue) {
            issues.Add(new PdfConversionProofIssue(
                "Xmp.PdfUaPart",
                options.RequiredXmpPdfUaPart.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "missing"));
        }
    }

    private static void AddTaggedContentIssues(List<PdfConversionProofIssue> issues, PdfDocumentInfo documentInfo, PdfConversionProofOptions options) {
        if (!HasRequiredTaggedContent(options)) {
            return;
        }

        PdfTaggedContentInfo? tagged = documentInfo.TaggedContent;
        if (tagged is null) {
            AddMissingTaggedContentIssues(issues, options);
            return;
        }

        for (int i = 0; i < options.RequiredTaggedStructureTypes.Count; i++) {
            string structureType = options.RequiredTaggedStructureTypes[i];
            if (!ContainsExact(tagged.StructureTypes, structureType)) {
                issues.Add(new PdfConversionProofIssue("TaggedContent.StructureType", structureType, "missing"));
            }
        }

        if (options.RequiredTaggedStructureElementCountAtLeast.HasValue &&
            tagged.StructureElementCount < options.RequiredTaggedStructureElementCountAtLeast.Value) {
            issues.Add(new PdfConversionProofIssue(
                "TaggedContent.StructureElementCount",
                "at least " + options.RequiredTaggedStructureElementCountAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                tagged.StructureElementCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        if (options.RequiredTaggedMarkedContentReferencesAtLeast.HasValue &&
            tagged.MarkedContentReferenceCount < options.RequiredTaggedMarkedContentReferencesAtLeast.Value) {
            issues.Add(new PdfConversionProofIssue(
                "TaggedContent.MarkedContentReferenceCount",
                "at least " + options.RequiredTaggedMarkedContentReferencesAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                tagged.MarkedContentReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
    }

    private static void AddMissingTaggedContentIssues(List<PdfConversionProofIssue> issues, PdfConversionProofOptions options) {
        for (int i = 0; i < options.RequiredTaggedStructureTypes.Count; i++) {
            issues.Add(new PdfConversionProofIssue("TaggedContent.StructureType", options.RequiredTaggedStructureTypes[i], "missing"));
        }

        if (options.RequiredTaggedStructureElementCountAtLeast.HasValue) {
            issues.Add(new PdfConversionProofIssue(
                "TaggedContent.StructureElementCount",
                "at least " + options.RequiredTaggedStructureElementCountAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "missing"));
        }

        if (options.RequiredTaggedMarkedContentReferencesAtLeast.HasValue) {
            issues.Add(new PdfConversionProofIssue(
                "TaggedContent.MarkedContentReferenceCount",
                "at least " + options.RequiredTaggedMarkedContentReferencesAtLeast.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "missing"));
        }
    }

    private static bool HasRequiredXmpMetadata(PdfConversionProofOptions options) {
        return options.RequiredXmpTitle is not null ||
            options.RequiredXmpCreator is not null ||
            options.RequiredXmpDescription is not null ||
            options.RequiredXmpProducer is not null ||
            options.RequiredXmpKeywords is not null ||
            options.RequiredXmpSubjects.Count > 0 ||
            options.RequiredXmpPdfAPart.HasValue ||
            options.RequiredXmpPdfAConformance is not null ||
            options.RequiredXmpPdfUaPart.HasValue;
    }

    private static bool HasRequiredTaggedContent(PdfConversionProofOptions options) {
        return options.RequiredTaggedStructureTypes.Count > 0 ||
            options.RequiredTaggedStructureElementCountAtLeast.HasValue ||
            options.RequiredTaggedMarkedContentReferencesAtLeast.HasValue;
    }
}
