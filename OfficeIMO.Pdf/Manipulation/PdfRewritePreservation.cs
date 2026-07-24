namespace OfficeIMO.Pdf;

/// <summary>
/// Provides reusable preservation checks for PDF rewrite and manipulation operations.
/// </summary>
public static partial class PdfRewritePreservation {
    private const double GeometryTolerance = 0.01d;

    /// <summary>
    /// Compares an original PDF and a rewritten PDF using the default preservation profile.
    /// </summary>
    public static PdfRewritePreservationReport Assess(byte[] originalPdf, byte[] rewrittenPdf) {
        return Assess(originalPdf, rewrittenPdf, null);
    }

    /// <summary>
    /// Compares an original PDF and a rewritten PDF using the supplied preservation profile.
    /// </summary>
    public static PdfRewritePreservationReport Assess(byte[] originalPdf, byte[] rewrittenPdf, PdfRewritePreservationOptions? options) {
        return Assess(originalPdf, rewrittenPdf, options, originalReadOptions: null, rewrittenReadOptions: null);
    }

    internal static PdfRewritePreservationReport Assess(
        byte[] originalPdf,
        byte[] rewrittenPdf,
        PdfRewritePreservationOptions? options,
        PdfReadOptions? originalReadOptions,
        PdfReadOptions? rewrittenReadOptions) {
        Guard.NotNull(originalPdf, nameof(originalPdf));
        Guard.NotNull(rewrittenPdf, nameof(rewrittenPdf));

        options ??= new PdfRewritePreservationOptions();
        PdfReadOptions? effectiveOriginalReadOptions = options.OriginalReadOptions ?? originalReadOptions;
        PdfReadOptions? effectiveRewrittenReadOptions = options.RewrittenReadOptions ?? rewrittenReadOptions;
        PdfDocumentInfo original = InspectReportInput(originalPdf, effectiveOriginalReadOptions, "original");
        PdfDocumentInfo rewritten = InspectReportInput(rewrittenPdf, effectiveRewrittenReadOptions, "rewritten");
        var issues = new List<PdfRewritePreservationIssue>();

        CompareCounts(issues, "PageCount", original.PageCount, rewritten.PageCount, options.PreservePageCount);
        ComparePageGeometry(issues, original, rewritten, options);
        CompareMetadata(issues, original.Metadata, rewritten.Metadata, options);
        CompareCounts(issues, "Outlines", original.Outlines.Count, rewritten.Outlines.Count, options.PreserveOutlines);
        CompareCounts(issues, "NamedDestinations", original.NamedDestinations.Count, rewritten.NamedDestinations.Count, options.PreserveNamedDestinations);
        CompareCounts(issues, "PageLabels", original.PageLabels.Count, rewritten.PageLabels.Count, options.PreservePageLabels);
        CompareCounts(issues, "LinkAnnotations", original.LinkAnnotationCount, rewritten.LinkAnnotationCount, options.PreserveLinkAnnotations);
        CompareCounts(issues, "Annotations", original.AnnotationCount, rewritten.AnnotationCount, options.PreserveAnnotations);
        CompareCounts(issues, "FormFields", original.FormFields.Count, rewritten.FormFields.Count, options.PreserveForms);
        CompareCounts(issues, "EmbeddedFiles", original.Attachments.Count, rewritten.Attachments.Count, options.PreserveEmbeddedFiles);
        CompareCounts(issues, "OutputIntents", original.OutputIntents.Count, rewritten.OutputIntents.Count, options.PreserveOutputIntents);
        CompareCounts(issues, "CatalogActions", original.CatalogActions.Count, rewritten.CatalogActions.Count, options.PreserveCatalogActions);
        CompareCounts(issues, "PageActions", original.Pages.Sum(static page => page.PageActions.Count), rewritten.Pages.Sum(static page => page.PageActions.Count), options.PreservePageActions);
        CompareNavigationMetadata(issues, original, rewritten, options);
        CompareViewerActionState(issues, original, rewritten, options);
        CompareBooleanMarker(issues, "Forms", original.HasForms, rewritten.HasForms, options.PreserveForms);
        CompareBooleanMarker(issues, "XmpMetadata", original.HasXmpMetadata, rewritten.HasXmpMetadata, options.PreserveXmpMetadata);
        CompareBooleanMarker(issues, "OptionalContent", original.HasOptionalContent, rewritten.HasOptionalContent, options.PreserveOptionalContent);
        CompareBooleanMarker(issues, "TaggedContent", original.HasTaggedContent, rewritten.HasTaggedContent, options.PreserveTaggedContent);
        CompareAttachments(issues, original.Attachments, rewritten.Attachments, options);
        CompareOutputIntents(issues, original.OutputIntents, rewritten.OutputIntents, options);
        CompareXmpMetadata(issues, original.XmpMetadata, rewritten.XmpMetadata, options);
        CompareOptionalContent(issues, original.OptionalContent, rewritten.OptionalContent, options);
        CompareTaggedContent(issues, original.TaggedContent, rewritten.TaggedContent, options);
        CompareSourceStructure(issues, original, rewritten, options);
        CompareSecurityState(issues, original.Security, rewritten.Security, options);
        CompareCatalogViewSettings(issues, original, rewritten, options);
        CompareTextMarkers(issues, rewrittenPdf, options.RequiredTextMarkers, effectiveRewrittenReadOptions);

        return new PdfRewritePreservationReport(original, rewritten, issues.AsReadOnly());
    }

    private static PdfDocumentInfo InspectReportInput(byte[] pdf, PdfReadOptions? readOptions, string inputName) {
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        document.DemandContentExtraction(inputName + " rewrite-preservation report");
        return PdfInspector.Inspect(pdf, document);
    }

    /// <summary>
    /// Compares two PDFs and throws when the rewritten PDF violates the default preservation profile.
    /// </summary>
    public static PdfRewritePreservationReport AssertPreserved(byte[] originalPdf, byte[] rewrittenPdf) {
        return AssertPreserved(originalPdf, rewrittenPdf, null);
    }

    /// <summary>
    /// Compares two PDFs and throws when the rewritten PDF violates the supplied preservation profile.
    /// </summary>
    public static PdfRewritePreservationReport AssertPreserved(byte[] originalPdf, byte[] rewrittenPdf, PdfRewritePreservationOptions? options) {
        PdfRewritePreservationReport report = Assess(originalPdf, rewrittenPdf, options);
        report.ThrowIfFailed();
        return report;
    }

    internal static PdfRewritePreservationReport AssertPreserved(
        byte[] originalPdf,
        byte[] rewrittenPdf,
        PdfRewritePreservationOptions? options,
        PdfReadOptions? originalReadOptions,
        PdfReadOptions? rewrittenReadOptions) {
        PdfRewritePreservationReport report = Assess(originalPdf, rewrittenPdf, options, originalReadOptions, rewrittenReadOptions);
        report.ThrowIfFailed();
        return report;
    }

    private static void CompareCounts(List<PdfRewritePreservationIssue> issues, string feature, int expected, int actual, bool enabled) {
        if (!enabled || expected == actual) {
            return;
        }

        issues.Add(CreateIssue(feature, expected.ToString(System.Globalization.CultureInfo.InvariantCulture), actual.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static void CompareBooleanMarker(List<PdfRewritePreservationIssue> issues, string feature, bool expected, bool actual, bool enabled) {
        if (!enabled || !expected || actual) {
            return;
        }

        issues.Add(CreateIssue(feature, "present", "missing"));
    }

    private static void ComparePageGeometry(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreservePageGeometry || original.PageCount != rewritten.PageCount) {
            return;
        }

        for (int i = 0; i < original.Pages.Count; i++) {
            PdfPageInfo before = original.Pages[i];
            PdfPageInfo after = rewritten.Pages[i];
            string prefix = "PageGeometry[" + before.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareDouble(issues, prefix + ".Width", before.Width, after.Width);
            CompareDouble(issues, prefix + ".Height", before.Height, after.Height);
            CompareCounts(issues, prefix + ".Rotation", before.RotationDegrees, after.RotationDegrees, true);
        }
    }

    private static void CompareMetadata(List<PdfRewritePreservationIssue> issues, PdfMetadata original, PdfMetadata rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveMetadata) {
            return;
        }

        CompareMetadataField(issues, options, "Title", original.Title, rewritten.Title);
        CompareMetadataField(issues, options, "Author", original.Author, rewritten.Author);
        CompareMetadataField(issues, options, "Subject", original.Subject, rewritten.Subject);
        CompareMetadataField(issues, options, "Keywords", original.Keywords, rewritten.Keywords);
    }

    private static void CompareMetadataField(List<PdfRewritePreservationIssue> issues, PdfRewritePreservationOptions options, string fieldName, string? expected, string? actual) {
        if (options.AllowedMetadataChanges.Contains(fieldName) || string.Equals(expected, actual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue("Metadata." + fieldName, expected ?? "(null)", actual ?? "(null)"));
    }

    private static void CompareAttachments(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfAttachmentInfo> original, IReadOnlyList<PdfAttachmentInfo> rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveEmbeddedFiles || original.Count != rewritten.Count) {
            return;
        }

        for (int i = 0; i < original.Count; i++) {
            PdfAttachmentInfo before = original[i];
            PdfAttachmentInfo after = rewritten[i];
            string prefix = "EmbeddedFiles[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Name", before.Name, after.Name);
            CompareString(issues, prefix + ".FileName", before.FileName, after.FileName);
            CompareString(issues, prefix + ".UnicodeFileName", before.UnicodeFileName, after.UnicodeFileName);
            CompareString(issues, prefix + ".Description", before.Description, after.Description);
            CompareString(issues, prefix + ".MimeType", before.MimeType, after.MimeType);
            CompareString(issues, prefix + ".Relationship", before.Relationship.ToString(), after.Relationship.ToString());
            CompareString(issues, prefix + ".Filter", before.Filter, after.Filter);
            CompareCounts(issues, prefix + ".SizeBytes", before.SizeBytes, after.SizeBytes, true);
            CompareNullableInt(issues, prefix + ".DeclaredSizeBytes", before.DeclaredSizeBytes, after.DeclaredSizeBytes);
            CompareString(issues, prefix + ".Source", before.Source, after.Source);
        }
    }

    private static void CompareOutputIntents(List<PdfRewritePreservationIssue> issues, IReadOnlyList<PdfOutputIntentInfo> original, IReadOnlyList<PdfOutputIntentInfo> rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveOutputIntents || original.Count != rewritten.Count) {
            return;
        }

        for (int i = 0; i < original.Count; i++) {
            PdfOutputIntentInfo before = original[i];
            PdfOutputIntentInfo after = rewritten[i];
            string prefix = "OutputIntents[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Subtype", before.Subtype, after.Subtype);
            CompareString(issues, prefix + ".OutputConditionIdentifier", before.OutputConditionIdentifier, after.OutputConditionIdentifier);
            CompareString(issues, prefix + ".OutputCondition", before.OutputCondition, after.OutputCondition);
            CompareString(issues, prefix + ".RegistryName", before.RegistryName, after.RegistryName);
            CompareString(issues, prefix + ".Info", before.Info, after.Info);
            CompareNullableInt(issues, prefix + ".DestinationOutputProfileColorComponents", before.DestinationOutputProfileColorComponents, after.DestinationOutputProfileColorComponents);
            CompareString(issues, prefix + ".DestinationOutputProfileAlternateColorSpace", before.DestinationOutputProfileAlternateColorSpace, after.DestinationOutputProfileAlternateColorSpace);
            CompareString(issues, prefix + ".DestinationOutputProfileFilter", before.DestinationOutputProfileFilter, after.DestinationOutputProfileFilter);
            CompareNullableInt(issues, prefix + ".DestinationOutputProfileSizeBytes", before.DestinationOutputProfileSizeBytes, after.DestinationOutputProfileSizeBytes);
            CompareNullableInt(issues, prefix + ".DestinationOutputProfileDeclaredSizeBytes", before.DestinationOutputProfileDeclaredSizeBytes, after.DestinationOutputProfileDeclaredSizeBytes);
            CompareString(issues, prefix + ".DestinationOutputProfileColorSpace", before.DestinationOutputProfileColorSpace, after.DestinationOutputProfileColorSpace);
            CompareNullableBoolean(issues, prefix + ".DestinationOutputProfileHasIccSignature", before.DestinationOutputProfileHasIccSignature, after.DestinationOutputProfileHasIccSignature);
        }
    }

    private static void CompareXmpMetadata(List<PdfRewritePreservationIssue> issues, PdfXmpMetadataInfo? original, PdfXmpMetadataInfo? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveXmpMetadata || original is null) {
            return;
        }

        if (rewritten is null) {
            issues.Add(CreateIssue("XmpMetadata.Readable", "present", "missing"));
            return;
        }

        CompareString(issues, "XmpMetadata.Subtype", original.Subtype, rewritten.Subtype);
        CompareString(issues, "XmpMetadata.Filter", original.Filter, rewritten.Filter);
        CompareCounts(issues, "XmpMetadata.StreamSizeBytes", original.StreamSizeBytes, rewritten.StreamSizeBytes, true);
        CompareCounts(issues, "XmpMetadata.DecodedSizeBytes", original.DecodedSizeBytes, rewritten.DecodedSizeBytes, true);
        CompareStringList(issues, "XmpMetadata.UnsupportedFilters", original.UnsupportedFilters, rewritten.UnsupportedFilters);
        CompareNullableBoolean(issues, "XmpMetadata.IsWellFormedXml", original.IsWellFormedXml, rewritten.IsWellFormedXml);
        CompareXmpMetadataField(issues, options, "Title", "Title", original.Title, rewritten.Title);
        CompareXmpMetadataField(issues, options, "Creator", "Author", original.Creator, rewritten.Creator);
        CompareXmpMetadataField(issues, options, "Description", "Subject", original.Description, rewritten.Description);
        CompareStringList(issues, "XmpMetadata.Subjects", original.Subjects, rewritten.Subjects, options.AllowedMetadataChanges.Contains("Keywords"));
        CompareString(issues, "XmpMetadata.Producer", original.Producer, rewritten.Producer);
        CompareXmpMetadataField(issues, options, "Keywords", "Keywords", original.Keywords, rewritten.Keywords);
        CompareNullableInt(issues, "XmpMetadata.PdfAPart", original.PdfAPart, rewritten.PdfAPart);
        CompareString(issues, "XmpMetadata.PdfAConformance", original.PdfAConformance, rewritten.PdfAConformance);
        CompareNullableInt(issues, "XmpMetadata.PdfUaPart", original.PdfUaPart, rewritten.PdfUaPart);
        CompareString(issues, "XmpMetadata.ElectronicInvoiceDocumentType", original.ElectronicInvoiceDocumentType, rewritten.ElectronicInvoiceDocumentType);
        CompareString(issues, "XmpMetadata.ElectronicInvoiceDocumentFileName", original.ElectronicInvoiceDocumentFileName, rewritten.ElectronicInvoiceDocumentFileName);
        CompareString(issues, "XmpMetadata.ElectronicInvoiceVersion", original.ElectronicInvoiceVersion, rewritten.ElectronicInvoiceVersion);
        CompareString(issues, "XmpMetadata.ElectronicInvoiceConformanceLevel", original.ElectronicInvoiceConformanceLevel, rewritten.ElectronicInvoiceConformanceLevel);
    }

    private static void CompareXmpMetadataField(List<PdfRewritePreservationIssue> issues, PdfRewritePreservationOptions options, string fieldName, string allowedMetadataFieldName, string? expected, string? actual) {
        if (options.AllowedMetadataChanges.Contains(allowedMetadataFieldName) || string.Equals(expected, actual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue("XmpMetadata." + fieldName, expected ?? "(null)", actual ?? "(null)"));
    }

    private static void CompareOptionalContent(List<PdfRewritePreservationIssue> issues, PdfOptionalContentProperties? original, PdfOptionalContentProperties? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveOptionalContent || original is null) {
            return;
        }

        if (rewritten is null) {
            issues.Add(CreateIssue("OptionalContent.Readable", "present", "missing"));
            return;
        }

        CompareString(issues, "OptionalContent.DefaultConfigurationName", original.DefaultConfigurationName, rewritten.DefaultConfigurationName);
        CompareString(issues, "OptionalContent.DefaultConfigurationCreator", original.DefaultConfigurationCreator, rewritten.DefaultConfigurationCreator);
        CompareString(issues, "OptionalContent.BaseState", original.BaseState, rewritten.BaseState);
        CompareOptionalContentReferences(issues, "OptionalContent.OnGroups", original.Groups, original.OnGroupObjectNumbers, rewritten.Groups, rewritten.OnGroupObjectNumbers);
        CompareOptionalContentReferences(issues, "OptionalContent.OffGroups", original.Groups, original.OffGroupObjectNumbers, rewritten.Groups, rewritten.OffGroupObjectNumbers);
        CompareOptionalContentReferences(issues, "OptionalContent.LockedGroups", original.Groups, original.LockedGroupObjectNumbers, rewritten.Groups, rewritten.LockedGroupObjectNumbers);
        CompareOptionalContentReferences(issues, "OptionalContent.OrderGroups", original.Groups, original.OrderGroupObjectNumbers, rewritten.Groups, rewritten.OrderGroupObjectNumbers);
        CompareCounts(issues, "OptionalContent.Groups", original.Groups.Count, rewritten.Groups.Count, true);

        int count = Math.Min(original.Groups.Count, rewritten.Groups.Count);
        for (int i = 0; i < count; i++) {
            PdfOptionalContentGroup before = original.Groups[i];
            PdfOptionalContentGroup after = rewritten.Groups[i];
            string prefix = "OptionalContent.Groups[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

            CompareString(issues, prefix + ".Name", before.Name, after.Name);
            CompareStringList(issues, prefix + ".Intents", before.Intents, after.Intents);
            CompareNullableBoolean(issues, prefix + ".IsInitiallyVisible", before.IsInitiallyVisible, after.IsInitiallyVisible);
            CompareNullableBoolean(issues, prefix + ".IsLocked", before.IsLocked, after.IsLocked);
            CompareNullableBoolean(issues, prefix + ".IsInDefaultOrder", before.IsInDefaultOrder, after.IsInDefaultOrder);
            CompareString(issues, prefix + ".ViewState", before.ViewState, after.ViewState);
            CompareString(issues, prefix + ".PrintState", before.PrintState, after.PrintState);
            CompareString(issues, prefix + ".ExportState", before.ExportState, after.ExportState);
            CompareString(issues, prefix + ".UsageCreator", before.UsageCreator, after.UsageCreator);
            CompareString(issues, prefix + ".UsageSubtype", before.UsageSubtype, after.UsageSubtype);
        }
    }

    private static void CompareOptionalContentReferences(
        List<PdfRewritePreservationIssue> issues,
        string feature,
        IReadOnlyList<PdfOptionalContentGroup> originalGroups,
        IReadOnlyList<int> originalObjectNumbers,
        IReadOnlyList<PdfOptionalContentGroup> rewrittenGroups,
        IReadOnlyList<int> rewrittenObjectNumbers) {
        string expected = FormatOptionalContentReferences(originalGroups, originalObjectNumbers);
        string actual = FormatOptionalContentReferences(rewrittenGroups, rewrittenObjectNumbers);
        if (string.Equals(expected, actual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue(feature, expected, actual));
    }

    private static string FormatOptionalContentReferences(IReadOnlyList<PdfOptionalContentGroup> groups, IReadOnlyList<int> objectNumbers) {
        if (objectNumbers.Count == 0) {
            return "(empty)";
        }

        var names = new List<string>(objectNumbers.Count);
        for (int i = 0; i < objectNumbers.Count; i++) {
            names.Add(GetOptionalContentReferenceName(groups, objectNumbers[i]));
        }

        return string.Join(",", names);
    }

    private static string GetOptionalContentReferenceName(IReadOnlyList<PdfOptionalContentGroup> groups, int objectNumber) {
        for (int i = 0; i < groups.Count; i++) {
            PdfOptionalContentGroup group = groups[i];
            if (group.ObjectNumber == objectNumber && !string.IsNullOrEmpty(group.Name)) {
                return group.Name;
            }
        }

        return "#" + objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static void CompareTaggedContent(List<PdfRewritePreservationIssue> issues, PdfTaggedContentInfo? original, PdfTaggedContentInfo? rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveTaggedContent || original is null) {
            return;
        }

        if (rewritten is null) {
            issues.Add(CreateIssue("TaggedContent.Readable", "present", "missing"));
            return;
        }

        CompareNullableBoolean(issues, "TaggedContent.Marked", original.Marked, rewritten.Marked);
        CompareNullableBoolean(issues, "TaggedContent.Suspects", original.Suspects, rewritten.Suspects);
        CompareNullableBoolean(issues, "TaggedContent.UserProperties", original.UserProperties, rewritten.UserProperties);
        CompareNullableInt(issues, "TaggedContent.ParentTreeNextKey", original.ParentTreeNextKey, rewritten.ParentTreeNextKey);
        CompareStringDictionary(issues, "TaggedContent.RoleMap", original.RoleMap, rewritten.RoleMap);
        CompareCounts(issues, "TaggedContent.RootElements", original.RootElementObjectNumbers.Count, rewritten.RootElementObjectNumbers.Count, true);
        CompareCounts(issues, "TaggedContent.ParentTreeEntries", original.ParentTreeEntryCount, rewritten.ParentTreeEntryCount, true);
        CompareCounts(issues, "TaggedContent.StructureElements", original.StructureElementCount, rewritten.StructureElementCount, true);
        CompareCounts(issues, "TaggedContent.MarkedContentReferences", original.MarkedContentReferenceCount, rewritten.MarkedContentReferenceCount, true);
        CompareCounts(issues, "TaggedContent.ObjectReferences", original.ObjectReferenceCount, rewritten.ObjectReferenceCount, true);
        CompareCounts(issues, "TaggedContent.LanguageElements", original.LanguageElementCount, rewritten.LanguageElementCount, true);
        CompareCounts(issues, "TaggedContent.AlternateTextElements", original.AlternateTextElementCount, rewritten.AlternateTextElementCount, true);
        CompareCounts(issues, "TaggedContent.FiguresWithoutAlternateText", original.FigureWithoutAlternateTextCount, rewritten.FigureWithoutAlternateTextCount, true);
        CompareStringList(issues, "TaggedContent.StructureTypes", original.StructureTypes, rewritten.StructureTypes);
        CompareStringIntDictionary(issues, "TaggedContent.StructureTypeCounts", original.StructureTypeCounts, rewritten.StructureTypeCounts);
        CompareNullableBoolean(issues, "TaggedContent.DeepEvidence", original.HasDeepTaggedPdfEvidence, rewritten.HasDeepTaggedPdfEvidence);
        CompareNullableBoolean(issues, "TaggedContent.FiguresHaveAlternateText", original.FiguresHaveAlternateText, rewritten.FiguresHaveAlternateText);

        int count = Math.Min(original.StructureElements.Count, rewritten.StructureElements.Count);
        for (int i = 0; i < count; i++) {
            CompareStructureElement(issues, original.StructureElements[i], rewritten.StructureElements[i], i);
        }
    }

    private static void CompareStructureElement(List<PdfRewritePreservationIssue> issues, PdfStructureElementInfo original, PdfStructureElementInfo rewritten, int index) {
        string prefix = "TaggedContent.StructureElements[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

        CompareString(issues, prefix + ".StructureType", original.StructureType, rewritten.StructureType);
        CompareString(issues, prefix + ".Language", original.Language, rewritten.Language);
        CompareString(issues, prefix + ".AlternateText", original.AlternateText, rewritten.AlternateText);
        CompareCounts(issues, prefix + ".ChildElements", original.ChildElementObjectNumbers.Count, rewritten.ChildElementObjectNumbers.Count, true);
        CompareCounts(issues, prefix + ".MarkedContentReferences", original.MarkedContentReferenceCount, rewritten.MarkedContentReferenceCount, true);
        CompareCounts(issues, prefix + ".ObjectReferences", original.ObjectReferenceCount, rewritten.ObjectReferenceCount, true);
    }

    private static void CompareSecurityState(List<PdfRewritePreservationIssue> issues, PdfDocumentSecurityInfo original, PdfDocumentSecurityInfo rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveSecurityState) {
            return;
        }

        CompareBooleanMarker(issues, "Security.Encryption", original.HasEncryption, rewritten.HasEncryption, true);
        CompareString(issues, "Security.EncryptionFilter", original.EncryptionFilter, rewritten.EncryptionFilter);
        CompareString(issues, "Security.EncryptionSubFilter", original.EncryptionSubFilter, rewritten.EncryptionSubFilter);
        CompareNullableInt(issues, "Security.EncryptionVersion", original.EncryptionVersion, rewritten.EncryptionVersion);
        CompareNullableInt(issues, "Security.EncryptionRevision", original.EncryptionRevision, rewritten.EncryptionRevision);
        CompareNullableInt(issues, "Security.EncryptionLengthBits", original.EncryptionLengthBits, rewritten.EncryptionLengthBits);
        CompareNullableInt(issues, "Security.EncryptionPermissions", original.EncryptionPermissions, rewritten.EncryptionPermissions);
        CompareNullableBoolean(issues, "Security.EncryptMetadata", original.EncryptMetadata, rewritten.EncryptMetadata);
        CompareBooleanMarker(issues, "Security.Signatures", original.HasSignatures, rewritten.HasSignatures, true);
        CompareCounts(issues, "Security.SignatureFields", original.SignatureFieldCount, rewritten.SignatureFieldCount, original.SignatureFieldCount > 0);
        CompareIntList(issues, "Security.SignatureFieldObjectNumbers", original.SignatureFieldObjectNumbers, rewritten.SignatureFieldObjectNumbers);
        CompareStringList(issues, "Security.SignatureFieldNames", original.SignatureFieldNames, rewritten.SignatureFieldNames);
        CompareCounts(issues, "Security.SignatureValues", original.SignatureValueCount, rewritten.SignatureValueCount, original.SignatureValueCount > 0);
        CompareCounts(issues, "Security.Signatures.Readable", original.Signatures.Count, rewritten.Signatures.Count, original.Signatures.Count > 0);
        CompareBooleanMarker(issues, "Security.ByteRange", original.HasByteRange, rewritten.HasByteRange, true);
        CompareCounts(issues, "Security.ByteRangeValues", original.ByteRangeValueCount, rewritten.ByteRangeValueCount, original.ByteRangeValueCount > 0);
        CompareNullableInt(issues, "Security.AcroFormSignatureFlags", original.AcroFormSignatureFlags, rewritten.AcroFormSignatureFlags);
        CompareBooleanMarker(issues, "Security.DocMDP", original.HasDocMDPPermissions, rewritten.HasDocMDPPermissions, true);
        CompareNullableInt(issues, "Security.DocMDPSignatureObjectNumber", original.DocMDPSignatureObjectNumber, rewritten.DocMDPSignatureObjectNumber);
        CompareString(issues, "Security.DocMDPTransformMethod", original.DocMDPTransformMethod, rewritten.DocMDPTransformMethod);
        CompareString(issues, "Security.DocMDPTransformVersion", original.DocMDPTransformVersion, rewritten.DocMDPTransformVersion);
        CompareNullableInt(issues, "Security.DocMDPPermissionLevel", original.DocMDPPermissionLevel, rewritten.DocMDPPermissionLevel);
        CompareBooleanMarker(issues, "Security.UsageRights", original.HasUsageRights, rewritten.HasUsageRights, true);
        CompareIntList(issues, "Security.UsageRightsObjectNumbers", original.UsageRightsObjectNumbers, rewritten.UsageRightsObjectNumbers);
        CompareDssInfo(issues, original.DocumentSecurityStore, rewritten.DocumentSecurityStore);
        CompareBooleanMarker(issues, "Security.TrailerId", original.HasTrailerId, rewritten.HasTrailerId, true);

        int count = Math.Min(original.Signatures.Count, rewritten.Signatures.Count);
        for (int i = 0; i < count; i++) {
            CompareSignatureInfo(issues, original.Signatures[i], rewritten.Signatures[i], i);
        }
    }

    private static void CompareSignatureInfo(List<PdfRewritePreservationIssue> issues, PdfSignatureInfo original, PdfSignatureInfo rewritten, int index) {
        string prefix = "Security.Signatures[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";

        CompareCounts(issues, prefix + ".ObjectNumber", original.ObjectNumber, rewritten.ObjectNumber, true);
        CompareNullableInt(issues, prefix + ".FieldObjectNumber", original.FieldObjectNumber, rewritten.FieldObjectNumber);
        CompareString(issues, prefix + ".FieldName", original.FieldName, rewritten.FieldName);
        CompareString(issues, prefix + ".Filter", original.Filter, rewritten.Filter);
        CompareString(issues, prefix + ".SubFilter", original.SubFilter, rewritten.SubFilter);
        CompareString(issues, prefix + ".SignerName", original.SignerName, rewritten.SignerName);
        CompareString(issues, prefix + ".Location", original.Location, rewritten.Location);
        CompareString(issues, prefix + ".Reason", original.Reason, rewritten.Reason);
        CompareString(issues, prefix + ".ContactInfo", original.ContactInfo, rewritten.ContactInfo);
        CompareString(issues, prefix + ".SigningTimeRaw", original.SigningTimeRaw, rewritten.SigningTimeRaw);
        CompareNullableBoolean(issues, prefix + ".HasByteRange", original.HasByteRange, rewritten.HasByteRange);
        CompareLongList(issues, prefix + ".ByteRangeValues", original.ByteRangeValues, rewritten.ByteRangeValues);
        CompareCounts(issues, prefix + ".ByteRangeValueCount", original.ByteRangeValueCount, rewritten.ByteRangeValueCount, true);
        CompareNullableBoolean(issues, prefix + ".HasContents", original.HasContents, rewritten.HasContents);
        CompareNullableInt(issues, prefix + ".ContentsSizeBytes", original.ContentsSizeBytes, rewritten.ContentsSizeBytes);
        CompareNullableInt(issues, prefix + ".ContentsEncodedSizeBytes", original.ContentsEncodedSizeBytes, rewritten.ContentsEncodedSizeBytes);
        CompareCounts(issues, prefix + ".ReferenceCount", original.ReferenceCount, rewritten.ReferenceCount, true);
        CompareSignatureFieldLock(issues, original.FieldLock, rewritten.FieldLock, prefix + ".FieldLock");
        CompareSignatureSeedValue(issues, original.SeedValue, rewritten.SeedValue, prefix + ".SeedValue");
    }

    private static void CompareSignatureFieldLock(List<PdfRewritePreservationIssue> issues, PdfSignatureFieldLockInfo? original, PdfSignatureFieldLockInfo? rewritten, string prefix) {
        if (original is null) {
            return;
        }

        if (rewritten is null) {
            issues.Add(CreateIssue(prefix, "present", "missing"));
            return;
        }

        CompareString(issues, prefix + ".Action", original.Action, rewritten.Action);
        CompareStringList(issues, prefix + ".Fields", original.Fields, rewritten.Fields);
    }

    private static void CompareSignatureSeedValue(List<PdfRewritePreservationIssue> issues, PdfSignatureSeedValueInfo? original, PdfSignatureSeedValueInfo? rewritten, string prefix) {
        if (original is null) {
            return;
        }

        if (rewritten is null) {
            issues.Add(CreateIssue(prefix, "present", "missing"));
            return;
        }

        CompareString(issues, prefix + ".Filter", original.Filter, rewritten.Filter);
        CompareStringList(issues, prefix + ".SubFilters", original.SubFilters, rewritten.SubFilters);
        CompareStringList(issues, prefix + ".DigestMethods", original.DigestMethods, rewritten.DigestMethods);
        CompareStringList(issues, prefix + ".Reasons", original.Reasons, rewritten.Reasons);
        CompareNullableInt(issues, prefix + ".Flags", original.Flags, rewritten.Flags);
        CompareNullableBoolean(issues, prefix + ".AddRevInfo", original.AddRevInfo, rewritten.AddRevInfo);
        CompareNullableInt(issues, prefix + ".MDPPermissionLevel", original.MDPPermissionLevel, rewritten.MDPPermissionLevel);
    }

    private static void CompareDssInfo(List<PdfRewritePreservationIssue> issues, PdfDocumentDssInfo original, PdfDocumentDssInfo rewritten) {
        CompareBooleanMarker(issues, "Security.DSS", original.HasDss, rewritten.HasDss, true);
        CompareNullableInt(issues, "Security.DSS.ObjectNumber", original.ObjectNumber, rewritten.ObjectNumber);
        CompareStringList(issues, "Security.DSS.VriKeys", original.VriKeys, rewritten.VriKeys);
        CompareIntList(issues, "Security.DSS.CertificateObjectNumbers", original.CertificateObjectNumbers, rewritten.CertificateObjectNumbers);
        CompareIntList(issues, "Security.DSS.OcspObjectNumbers", original.OcspObjectNumbers, rewritten.OcspObjectNumbers);
        CompareIntList(issues, "Security.DSS.CrlObjectNumbers", original.CrlObjectNumbers, rewritten.CrlObjectNumbers);
        CompareIntList(issues, "Security.DSS.VriCertificateObjectNumbers", original.VriCertificateObjectNumbers, rewritten.VriCertificateObjectNumbers);
        CompareIntList(issues, "Security.DSS.VriOcspObjectNumbers", original.VriOcspObjectNumbers, rewritten.VriOcspObjectNumbers);
        CompareIntList(issues, "Security.DSS.VriCrlObjectNumbers", original.VriCrlObjectNumbers, rewritten.VriCrlObjectNumbers);
        CompareIntList(issues, "Security.DSS.TimestampObjectNumbers", original.TimestampObjectNumbers, rewritten.TimestampObjectNumbers);
    }

    private static void CompareCatalogViewSettings(List<PdfRewritePreservationIssue> issues, PdfDocumentInfo original, PdfDocumentInfo rewritten, PdfRewritePreservationOptions options) {
        if (!options.PreserveCatalogViewSettings) {
            return;
        }

        CompareString(issues, "CatalogPageMode", original.CatalogPageMode, rewritten.CatalogPageMode);
        CompareString(issues, "CatalogPageLayout", original.CatalogPageLayout, rewritten.CatalogPageLayout);
        CompareString(issues, "CatalogLanguage", original.CatalogLanguage, rewritten.CatalogLanguage);
        CompareBooleanMarker(issues, "ViewerPreferences", original.HasViewerPreferences, rewritten.HasViewerPreferences, options.PreserveViewerPreferences);
    }

    private static void CompareTextMarkers(
        List<PdfRewritePreservationIssue> issues,
        byte[] rewrittenPdf,
        IEnumerable<string> requiredTextMarkers,
        PdfReadOptions? readOptions) {
        string text = string.Empty;
        bool loaded = false;

        foreach (string marker in requiredTextMarkers) {
            if (string.IsNullOrEmpty(marker)) {
                continue;
            }

            if (!loaded) {
                text = PdfReadDocument.Open(rewrittenPdf, readOptions).ExtractText();
                loaded = true;
            }

            if (text.IndexOf(marker, StringComparison.Ordinal) < 0) {
                issues.Add(CreateIssue("TextMarker", marker, "missing"));
            }
        }
    }

    private static void CompareDouble(List<PdfRewritePreservationIssue> issues, string feature, double expected, double actual) {
        if (Math.Abs(expected - actual) <= GeometryTolerance) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatDouble(expected), FormatDouble(actual)));
    }

    private static void CompareString(List<PdfRewritePreservationIssue> issues, string feature, string? expected, string? actual) {
        if (string.Equals(expected, actual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue(feature, expected ?? "(null)", actual ?? "(null)"));
    }

    private static void CompareStringList(List<PdfRewritePreservationIssue> issues, string feature, IReadOnlyList<string> expected, IReadOnlyList<string> actual, bool skip = false) {
        if (skip || expected.SequenceEqual(actual, StringComparer.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatStringList(expected), FormatStringList(actual)));
    }

    private static void CompareIntList(List<PdfRewritePreservationIssue> issues, string feature, IReadOnlyList<int> expected, IReadOnlyList<int> actual) {
        if (expected.SequenceEqual(actual)) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatIntList(expected), FormatIntList(actual)));
    }

    private static void CompareStringDictionary(List<PdfRewritePreservationIssue> issues, string feature, IReadOnlyDictionary<string, string> expected, IReadOnlyDictionary<string, string> actual) {
        string formattedExpected = FormatStringDictionary(expected);
        string formattedActual = FormatStringDictionary(actual);
        if (string.Equals(formattedExpected, formattedActual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue(feature, formattedExpected, formattedActual));
    }

    private static void CompareStringIntDictionary(List<PdfRewritePreservationIssue> issues, string feature, IReadOnlyDictionary<string, int> expected, IReadOnlyDictionary<string, int> actual) {
        string formattedExpected = FormatStringIntDictionary(expected);
        string formattedActual = FormatStringIntDictionary(actual);
        if (string.Equals(formattedExpected, formattedActual, StringComparison.Ordinal)) {
            return;
        }

        issues.Add(CreateIssue(feature, formattedExpected, formattedActual));
    }

    private static void CompareLongList(List<PdfRewritePreservationIssue> issues, string feature, IReadOnlyList<long> expected, IReadOnlyList<long> actual) {
        if (expected.SequenceEqual(actual)) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatLongList(expected), FormatLongList(actual)));
    }

    private static void CompareNullableInt(List<PdfRewritePreservationIssue> issues, string feature, int? expected, int? actual) {
        if (expected == actual) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatNullable(expected), FormatNullable(actual)));
    }

    private static void CompareMinimumCount(List<PdfRewritePreservationIssue> issues, string feature, int expectedMinimum, int actual, bool enabled) {
        if (!enabled || actual >= expectedMinimum) {
            return;
        }

        issues.Add(CreateIssue(
            feature,
            "at least " + expectedMinimum.ToString(System.Globalization.CultureInfo.InvariantCulture),
            actual.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static void CompareNullableBoolean(List<PdfRewritePreservationIssue> issues, string feature, bool? expected, bool? actual) {
        if (expected == actual) {
            return;
        }

        issues.Add(CreateIssue(feature, FormatNullable(expected), FormatNullable(actual)));
    }

    private static PdfRewritePreservationIssue CreateIssue(string feature, string expected, string actual) {
        return new PdfRewritePreservationIssue(
            feature,
            expected,
            actual,
            feature + " expected " + expected + " but was " + actual + ".");
    }

    private static string FormatDouble(double value) {
        return value.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string FormatStringList(IReadOnlyList<string> values) {
        return values.Count == 0 ? "(empty)" : string.Join(",", values);
    }

    private static string FormatIntList(IReadOnlyList<int> values) {
        return values.Count == 0 ? "(empty)" : string.Join(",", values.Select(static value => value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static string FormatLongList(IReadOnlyList<long> values) {
        return values.Count == 0 ? "(empty)" : string.Join(",", values.Select(static value => value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static string FormatStringDictionary(IReadOnlyDictionary<string, string> values) {
        if (values.Count == 0) {
            return "(empty)";
        }

        return string.Join(
            ",",
            values
                .OrderBy(static value => value.Key, StringComparer.Ordinal)
                .Select(static value => value.Key + "=" + value.Value));
    }

    private static string FormatStringIntDictionary(IReadOnlyDictionary<string, int> values) {
        if (values.Count == 0) {
            return "(empty)";
        }

        return string.Join(
            ",",
            values
                .OrderBy(static value => value.Key, StringComparer.Ordinal)
                .Select(static value => value.Key + "=" + value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    private static string FormatNullable(int? value) {
        return value.HasValue
            ? value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : "(null)";
    }

    private static string FormatNullable(bool? value) {
        return value.HasValue
            ? value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : "(null)";
    }
}
