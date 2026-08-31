using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal static class BrowserPdfPolicy {
    internal const long MaxInputBytes = 25L * 1024L * 1024L;
    internal const int MaxPages = 500;
    internal const int MaxSplitDocuments = 100;
    internal const long MaxOutputBytes = 96L * 1024L * 1024L;
    internal const long MaxSplitSerializedBytes = 64L * 1024L * 1024L;

    internal static PdfDocument Open(SelectedDocument file, string? password = null) =>
        PdfDocument.Load(file.Bytes, CreateReadOptions(password));

    internal static PdfLoadOptions CreateReadOptions(string? password = null) => new() {
        Password = password,
        AesCryptographyProvider = OfficeManagedAesCryptographyProvider.Default,
        Limits = new PdfReadLimits {
            MaxInputBytes = MaxInputBytes,
            MaxIndirectObjects = 50_000,
            MaxRawStreamBytes = 32 * 1024 * 1024,
            MaxDecodedStreamBytes = 32 * 1024 * 1024,
            MaxTotalDecodedStreamBytes = 96L * 1024L * 1024L,
            MaxPageContentBytes = 32 * 1024 * 1024,
            MaxRetainedContentBytes = 96L * 1024L * 1024L,
            MaxActualTextCharacters = 250_000,
            MaxDecodedTextCharacters = 1_000_000,
            MaxTextSearchMatches = 10_000,
            MaxObjectCharacters = 250_000,
            MaxTokensPerObject = 50_000,
            MaxObjectNestingDepth = 64,
            MaxObjectParsingTime = TimeSpan.FromSeconds(10),
            MaxRevisions = 1_000,
            MaxPageTreeNodes = 2_000,
            MaxPageTreeDepth = 128,
            MaxPages = BrowserPdfPolicy.MaxPages,
            MaxFormFields = 10_000,
            MaxFormFieldDepth = 128,
            MaxNameTreeNodes = 10_000,
            MaxNameTreeDepth = 64,
            MaxJavaScriptBytes = 1_000_000,
            MaxJavaScripts = 100,
            MaxWidgetActions = 10_000,
            MaxTotalJavaScriptBytes = 4L * 1024L * 1024L,
            MaxAttachments = 1_000,
            MaxTotalAttachmentBytes = 32L * 1024L * 1024L,
            MaxFormFieldAppearanceStates = 1_024,
            MaxAnnotationsPerPage = 10_000,
            MaxColorSpaceResourcesPerPage = 1_024,
            MaxContentOperations = 250_000,
            MaxContentOperands = 500_000,
            MaxContentNestingDepth = 64,
            MaxType3GlyphInvocationsPerPage = 100_000
        }
    };
}
