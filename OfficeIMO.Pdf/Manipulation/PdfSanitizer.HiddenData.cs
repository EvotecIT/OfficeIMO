namespace OfficeIMO.Pdf;

internal static partial class PdfSanitizer {
    private static readonly string[] UserMetadataKeys = { "Title", "Author", "Subject", "Keywords", "Creator" };

    private static PdfSanitizationReport BuildReport(
        byte[] pdf,
        PdfSanitizationOptions policy,
        PdfLoadOptions? readOptions,
        IReadOnlyList<PdfSanitizationFinding> findings) {
        policy.CancellationToken.ThrowIfCancellationRequested();
        PdfDocumentInfo info = PdfInspector.Inspect(pdf, readOptions);
        int userMetadata = policy.ShouldRemoveUserMetadata
            ? CountUserMetadataEntries(pdf, readOptions, policy.CancellationToken)
            : 0;
        int embeddedFiles = policy.ShouldRemoveEmbeddedFiles ? info.Attachments.Count : 0;
        int commentsAndMarkup = 0;
        for (int i = 0; i < info.Annotations.Count; i++) {
            string subtype = info.Annotations[i].Subtype;
            if (policy.ShouldRemoveCommentAnnotation(subtype) || policy.ShouldRemoveLegacyRichAnnotation(subtype)) {
                commentsAndMarkup++;
            }
        }
        int bookmarks = policy.ShouldRemoveBookmarks ? CountOutlines(info.Outlines) : 0;
        int optionalContent = policy.ShouldRemoveOptionalContent && info.HasOptionalContent
            ? Math.Max(1, info.OptionalContentGroupCount)
            : 0;
        return new PdfSanitizationReport(
            findings,
            userMetadata,
            embeddedFiles,
            commentsAndMarkup,
            bookmarks,
            optionalContent);
    }

    private static int CountUserMetadataEntries(
        byte[] pdf,
        PdfLoadOptions? readOptions,
        System.Threading.CancellationToken cancellationToken) {
        var parsed = PdfSyntax.ParseObjects(pdf, readOptions, out _, out _, cancellationToken);
        PdfDocumentSecurityInfo baseline = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            readOptions,
            includeParsedDetails: false,
            cancellationToken: cancellationToken);
        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(
            pdf,
            parsed.Map,
            parsed.TrailerRaw,
            baseline,
            readOptions,
            cancellationToken);
        int count = 0;
        if (security.InfoObjectNumber.HasValue &&
            parsed.Map.TryGetValue(security.InfoObjectNumber.Value, out PdfIndirectObject? infoObject) &&
            infoObject.Value is PdfDictionary info) {
            for (int i = 0; i < UserMetadataKeys.Length; i++) {
                if (info.Items.ContainsKey(UserMetadataKeys[i])) count++;
            }
        }
        if (security.RootObjectNumber.HasValue &&
            parsed.Map.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) &&
            rootObject.Value is PdfDictionary catalog &&
            catalog.Items.ContainsKey("Metadata")) {
            count++;
        }
        return count;
    }

    private static int CountOutlines(IReadOnlyList<PdfOutlineItem> outlines) {
        int count = 0;
        for (int i = 0; i < outlines.Count; i++) {
            count = checked(count + 1 + CountOutlines(outlines[i].Children));
        }
        return count;
    }

    private static void SanitizeDocumentContainers(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        PdfSanitizationOptions policy) {
        if (policy.ShouldRemoveUserMetadata && security.InfoObjectNumber.HasValue &&
            objects.TryGetValue(security.InfoObjectNumber.Value, out PdfIndirectObject? infoObject) &&
            infoObject.Value is PdfDictionary info) {
            for (int i = 0; i < UserMetadataKeys.Length; i++) info.Items.Remove(UserMetadataKeys[i]);
        }

        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog) {
            throw new InvalidOperationException("The active PDF trailer does not reference a readable catalog object.");
        }

        if (policy.ShouldRemoveUserMetadata) catalog.Items.Remove("Metadata");
        if (policy.ShouldRemoveBookmarks) {
            catalog.Items.Remove("Outlines");
            if (catalog.Get<PdfName>("PageMode")?.Name == "UseOutlines") catalog.Items.Remove("PageMode");
        }
        if (policy.ShouldRemoveOptionalContent) catalog.Items.Remove("OCProperties");
    }

    private static bool ShouldRemoveAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        PdfSanitizationOptions policy) {
        if (Resolve(objects, annotation.Get<PdfObject>("Subtype")) is not PdfName subtype) return false;
        if (string.Equals(subtype.Name, "FileAttachment", StringComparison.Ordinal) && policy.ShouldRemoveEmbeddedFiles) return true;
        return policy.ShouldRemoveCommentAnnotation(subtype.Name) || policy.ShouldRemoveLegacyRichAnnotation(subtype.Name);
    }

    private static void RemoveOptionalContentAssociation(PdfDictionary dictionary) {
        string? type = dictionary.Get<PdfName>("Type")?.Name;
        if (string.Equals(type, "OCG", StringComparison.Ordinal) || string.Equals(type, "OCMD", StringComparison.Ordinal)) {
            dictionary.Items.Clear();
            return;
        }
        dictionary.Items.Remove("OC");
        dictionary.Items.Remove("OCGs");
    }
}
