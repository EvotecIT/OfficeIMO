using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal sealed class SerializationContext {
        public SerializationContext(
            Dictionary<int, int> numberMap,
            int pagesObjectId,
            Dictionary<int, Dictionary<string, PdfObject>> materializedPageValues,
            Dictionary<int, PdfIndirectObject>? sourceObjects = null,
            Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null,
            bool preserveReferenceGenerations = false,
            bool preserveRawStringBytes = false) {
            NumberMap = numberMap;
            PagesObjectId = pagesObjectId;
            MaterializedPageValues = materializedPageValues;
            SourceObjectGenerations = sourceObjects?.ToDictionary(entry => entry.Key, entry => entry.Value.Generation) ?? new Dictionary<int, int>();
            PageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
            PreserveReferenceGenerations = preserveReferenceGenerations;
            PreserveRawStringBytes = preserveRawStringBytes;
        }
    
        public Dictionary<int, int> NumberMap { get; }
    
        public int PagesObjectId { get; }
    
        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; }
    
        public Dictionary<int, int> SourceObjectGenerations { get; }

        public bool PreserveReferenceGenerations { get; }

        public bool PreserveRawStringBytes { get; }
    
        public Dictionary<int, Dictionary<string, PdfObject>> PageOverrides { get; }
    }
    
    internal sealed class AdditionalObject {
        public AdditionalObject(int pseudoObjectNumber, PdfObject value) {
            PseudoObjectNumber = pseudoObjectNumber;
            Value = value;
        }
    
        public int PseudoObjectNumber { get; }
    
        public PdfObject Value { get; }
    }
    
    private sealed class ClonedPageObject {
        public ClonedPageObject(
            int sourcePageObjectNumber,
            int outputPageObjectNumber,
            Dictionary<string, PdfObject>? pageOverrides,
            Dictionary<int, int> annotationObjectMap) {
            SourcePageObjectNumber = sourcePageObjectNumber;
            OutputPageObjectNumber = outputPageObjectNumber;
            PageOverrides = pageOverrides;
            AnnotationObjectMap = annotationObjectMap;
        }
    
        public int SourcePageObjectNumber { get; }
    
        public int OutputPageObjectNumber { get; }
    
        public Dictionary<string, PdfObject>? PageOverrides { get; }
    
        public Dictionary<int, int> AnnotationObjectMap { get; }
    }
    
    private sealed class ClonedAnnotationState {
        public static readonly ClonedAnnotationState Empty = new ClonedAnnotationState(null, new Dictionary<int, int>());
    
        public ClonedAnnotationState(Dictionary<string, PdfObject>? pageOverrides, Dictionary<int, int> annotationObjectMap) {
            PageOverrides = pageOverrides;
            AnnotationObjectMap = annotationObjectMap;
        }
    
        public Dictionary<string, PdfObject>? PageOverrides { get; }
    
        public Dictionary<int, int> AnnotationObjectMap { get; }
    }
    
    private sealed class PageLabelEntry {
        public PageLabelEntry(int startPageIndex, PdfDictionary labelDictionary) {
            StartPageIndex = startPageIndex;
            LabelDictionary = labelDictionary;
        }
    
        public int StartPageIndex { get; }
    
        public PdfDictionary LabelDictionary { get; }
    }
    
    private sealed class NamedDestinationNameTreeEntry {
        public NamedDestinationNameTreeEntry(PdfStringObj name, PdfObject destination) {
            Name = name;
            Destination = destination;
        }
    
        public PdfStringObj Name { get; }
    
        public PdfObject Destination { get; }
    }
    
    internal sealed class CatalogRewriteState {
        public static readonly CatalogRewriteState Empty = new CatalogRewriteState(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);
    
        public CatalogRewriteState(string? pageMode, string? pageLayout, PdfObject? catalogVersion, PdfObject? catalogLanguage, PdfObject? outlines, PdfObject? pageLabels, PdfObject? namedDestinations, PdfObject? namedDestinationNameTree, PdfObject? openAction, PdfObject? viewerPreferences, PdfObject? xmpMetadata, PdfObject? catalogUri, PdfObject? outputIntents, PdfObject? embeddedFiles, PdfObject? associatedFiles, PdfObject? optionalContent, IReadOnlyList<int>? sourcePageObjectNumbers = null) {
            PageMode = string.IsNullOrEmpty(pageMode) ? null : pageMode;
            PageLayout = string.IsNullOrEmpty(pageLayout) ? null : pageLayout;
            CatalogVersion = catalogVersion;
            CatalogLanguage = catalogLanguage;
            Outlines = outlines;
            PageLabels = pageLabels;
            NamedDestinations = namedDestinations;
            NamedDestinationNameTree = namedDestinationNameTree;
            OpenAction = openAction;
            ViewerPreferences = viewerPreferences;
            XmpMetadata = xmpMetadata;
            CatalogUri = catalogUri;
            OutputIntents = outputIntents;
            EmbeddedFiles = embeddedFiles;
            AssociatedFiles = associatedFiles;
            OptionalContent = optionalContent;
            SourcePageObjectNumbers = sourcePageObjectNumbers;
        }
    
        public string? PageMode { get; }
    
        public string? PageLayout { get; }
    
        public PdfObject? CatalogVersion { get; }
    
        public PdfObject? CatalogLanguage { get; }
    
        public PdfObject? Outlines { get; }
    
        public PdfObject? PageLabels { get; }
    
        public PdfObject? NamedDestinations { get; }
    
        public PdfObject? NamedDestinationNameTree { get; }
    
        public PdfObject? OpenAction { get; }
    
        public PdfObject? ViewerPreferences { get; }
    
        public PdfObject? XmpMetadata { get; }
    
        public PdfObject? CatalogUri { get; }
    
        public PdfObject? OutputIntents { get; }
    
        public PdfObject? EmbeddedFiles { get; }
    
        public PdfObject? AssociatedFiles { get; }
    
        public PdfObject? OptionalContent { get; }
    
        public IReadOnlyList<int>? SourcePageObjectNumbers { get; }
    }
}
