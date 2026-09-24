using System.Globalization;
using System.Threading;

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
            bool preserveRawStringBytes = false,
            CancellationToken cancellationToken = default) {
            NumberMap = numberMap;
            PagesObjectId = pagesObjectId;
            MaterializedPageValues = materializedPageValues;
            SourceObjects = sourceObjects;
            PageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
            PreserveReferenceGenerations = preserveReferenceGenerations;
            PreserveRawStringBytes = preserveRawStringBytes;
            CancellationToken = cancellationToken;
        }
    
        public Dictionary<int, int> NumberMap { get; }
    
        public int PagesObjectId { get; }
    
        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; }
    
        /// <summary>Uses the existing parse for reference-generation checks without copying every source object.</summary>
        public Dictionary<int, PdfIndirectObject>? SourceObjects { get; }

        public bool PreserveReferenceGenerations { get; }

        public bool PreserveRawStringBytes { get; }

        public CancellationToken CancellationToken { get; }
    
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
    
    internal sealed class PageLabelEntry {
        public PageLabelEntry(int startPageIndex, PdfDictionary labelDictionary) {
            StartPageIndex = startPageIndex;
            LabelDictionary = labelDictionary;
        }
    
        public int StartPageIndex { get; }
    
        public PdfDictionary LabelDictionary { get; }
    }
    
    internal sealed class NamedDestinationNameTreeEntry {
        public NamedDestinationNameTreeEntry(PdfStringObj name, PdfObject destination, int order = 0) {
            Name = name;
            Destination = destination;
            Order = order;
        }
    
        public PdfStringObj Name { get; }
    
        public PdfObject Destination { get; }

        public int Order { get; }
    }
    
    internal sealed class CatalogRewriteState {
        private const int MinimumIndexedDestinationCount = 128;
        public static readonly CatalogRewriteState Empty = new CatalogRewriteState(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        private readonly Lazy<Dictionary<int, int>>? _sourcePageIndexes;
        private readonly Lazy<List<PageLabelEntry>?>? _pageLabelEntries;
        private readonly Lazy<Dictionary<int, List<NamedDestinationNameTreeEntry>>?>? _namedDestinationPageIndex;
        private readonly Lazy<Dictionary<int, List<DirectNamedDestinationEntry>>?>? _directNamedDestinationPageIndex;
        private int _namedDestinationFilterCount;
        private int _directNamedDestinationFilterCount;
    
        public CatalogRewriteState(string? pageMode, string? pageLayout, PdfObject? catalogVersion, PdfObject? catalogLanguage, PdfObject? outlines, PdfObject? pageLabels, PdfObject? namedDestinations, PdfObject? namedDestinationNameTree, PdfObject? openAction, PdfObject? viewerPreferences, PdfObject? xmpMetadata, PdfObject? catalogUri, PdfObject? outputIntents, PdfObject? embeddedFiles, PdfObject? associatedFiles, PdfObject? optionalContent, List<int>? sourcePageObjectNumbers = null, Dictionary<int, PdfIndirectObject>? sourceObjects = null, CancellationToken cancellationToken = default) {
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
            if (pageLabels is not null && sourcePageObjectNumbers is not null) {
                _sourcePageIndexes = new Lazy<Dictionary<int, int>>(
                    () => BuildSourcePageIndexes(sourcePageObjectNumbers, cancellationToken));
            }
            if (pageLabels is not null && sourceObjects is not null) {
                _pageLabelEntries = new Lazy<List<PageLabelEntry>?>(
                    () => ReadPageLabelEntries(sourceObjects, pageLabels, cancellationToken));
            }
            if (namedDestinationNameTree is not null && sourceObjects is not null) {
                _namedDestinationPageIndex = new Lazy<Dictionary<int, List<NamedDestinationNameTreeEntry>>?>(
                    () => BuildNamedDestinationPageIndex(sourceObjects, namedDestinationNameTree, cancellationToken));
            }
            if (namedDestinations is not null &&
                sourceObjects is not null &&
                ResolveDictionary(sourceObjects, namedDestinations) is { Items.Count: >= MinimumIndexedDestinationCount }) {
                _directNamedDestinationPageIndex = new Lazy<Dictionary<int, List<DirectNamedDestinationEntry>>?>(
                    () => BuildDirectNamedDestinationPageIndex(sourceObjects, namedDestinations, cancellationToken));
            }
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
    
        public List<int>? SourcePageObjectNumbers { get; }

        /// <summary>Shares the source page lookup across outputs of a compound extraction.</summary>
        internal Dictionary<int, int>? SourcePageIndexes => _sourcePageIndexes?.Value;

        /// <summary>Shares parsed label rules across outputs of a compound extraction.</summary>
        internal List<PageLabelEntry>? PageLabelEntries => _pageLabelEntries?.Value;

        /// <summary>Builds the page index only when a catalog is filtered for multiple outputs.</summary>
        internal Dictionary<int, List<NamedDestinationNameTreeEntry>>? GetNamedDestinationPageIndexForRepeatedUse() {
            if (_namedDestinationPageIndex is null ||
                System.Threading.Interlocked.Increment(ref _namedDestinationFilterCount) == 1) {
                return null;
            }

            return _namedDestinationPageIndex.Value;
        }

        /// <summary>Indexes direct destinations only after a second extraction uses this source.</summary>
        internal Dictionary<int, List<DirectNamedDestinationEntry>>? GetDirectNamedDestinationPageIndexForRepeatedUse() {
            if (_directNamedDestinationPageIndex is null ||
                System.Threading.Interlocked.Increment(ref _directNamedDestinationFilterCount) == 1) {
                return null;
            }

            return _directNamedDestinationPageIndex.Value;
        }
    }

    internal readonly struct DirectNamedDestinationEntry {
        internal DirectNamedDestinationEntry(string name, PdfObject destination, int order) {
            Name = name;
            Destination = destination;
            Order = order;
        }

        internal string Name { get; }
        internal PdfObject Destination { get; }
        internal int Order { get; }
    }
}
