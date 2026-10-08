using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext : IDisposable {
        private StringBuilder sb = new StringBuilder();
        private readonly PdfPageContentStore pageContents;
        private readonly bool ownsPageContents;
        private readonly bool isRunningContent;
        private bool pageContentsTransferred;
        private readonly System.Collections.Generic.List<LayoutResult.Page> pages = new System.Collections.Generic.List<LayoutResult.Page>();
        private readonly System.Collections.Generic.Stack<PdfOptions> optionsStack = new System.Collections.Generic.Stack<PdfOptions>();
        private readonly System.Collections.Generic.Stack<int> pageGroupStack = new System.Collections.Generic.Stack<int>();
        private readonly System.Collections.Generic.HashSet<string> emittedTableCellNamedDestinations = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
        private readonly bool emitGeneratedStructure;
        private readonly System.Collections.Generic.IReadOnlyList<SectionBlock> sectionDefinitions;
        private readonly System.Collections.Generic.IReadOnlyDictionary<string, int> sectionPageNumbers;
        private readonly System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>> deferredMaterializations;
        private readonly System.Threading.CancellationToken cancellationToken;
        private readonly int? maximumGeneratedPages;
        private long startedPageCount;
        private readonly System.Collections.Generic.List<SectionBlock> encounteredSectionDefinitions = new System.Collections.Generic.List<SectionBlock>();
        private readonly System.Collections.Generic.Dictionary<System.Collections.Generic.List<ColItem>, double[]> rowColumnKeepChainHeights = new System.Collections.Generic.Dictionary<System.Collections.Generic.List<ColItem>, double[]>();
        private bool encounteredTableOfContents;
        private PdfOptions currentOpts;
        private PdfOptions currentPageBaseOptions;
        private readonly Dictionary<PdfOptions, PdfOptions> mirroredPageOptions = new();
        private readonly HashSet<int> emittedPageGroups = new();
        private readonly Dictionary<int, int> emittedPageGroupCounts = new();
        private int previousVisiblePageNumber;
        private int currentVisiblePageNumber;
        private int currentPageGroupId;
        private int nextPageGroupId = 1;
        private LayoutResult.Page? currentPage;
        private double width;
        private double yStart;
        private double y;
        private bool pageDirty;
        private bool usedBold;
        private bool usedItalic;
        private bool usedBoldItalic;
        private int _canvasClipDepth;
        private bool _suppressCanvasAccessibilityWrappers;
        private bool _suppressCanvasStructureRegistration;
        private bool _suppressCanvasActualTextChildren;
        private PageStructElement? _canvasStructureParentElement;
        private readonly System.Collections.Generic.List<FlowSemanticScope> flowSemanticScopes = new System.Collections.Generic.List<FlowSemanticScope>();
        private bool stopDocumentFlow;
        private readonly System.Collections.Generic.HashSet<PdfLayoutPositionCapture> initializedPositionCaptures = new System.Collections.Generic.HashSet<PdfLayoutPositionCapture>();
        private readonly System.Collections.Generic.List<PdfLayerDefinition> activeLayers = new System.Collections.Generic.List<PdfLayerDefinition>();
        private readonly System.Collections.Generic.List<ContainerRenderScope> activeContainerScopes = new System.Collections.Generic.List<ContainerRenderScope>();
        private readonly System.Collections.Generic.Dictionary<(string Key, string Type, PageStructElement? Parent, string Scope, int Columns, int Rows, string AlternativeText), PageStructElement> canvasStructureElements = new System.Collections.Generic.Dictionary<(string, string, PageStructElement?, string, int, int, string), PageStructElement>();

        public LayoutContext(
            PdfOptions options,
            System.Collections.Generic.IReadOnlyList<SectionBlock>? sections = null,
            System.Collections.Generic.IReadOnlyDictionary<string, int>? resolvedSectionPages = null,
            System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>>? materializations = null,
            PdfPageContentStore? sharedPageContents = null,
            bool isRunningContent = false,
            RunningContentImageAssets? runningAssets = null,
            IReadOnlyList<PageNumberInfo>? previousRunningPages = null,
            int previousDocumentPages = 1,
            System.Threading.CancellationToken cancellationToken = default) {
            currentOpts = options;
            currentPageBaseOptions = options;
            this.cancellationToken = cancellationToken;
            maximumGeneratedPages = options.MaxGeneratedPages;
            pageContents = sharedPageContents ?? new PdfPageContentStore(options.PageContentMemoryLimitBytes);
            ownsPageContents = sharedPageContents == null;
            this.isRunningContent = isRunningContent;
            this.runningAssets = runningAssets ?? new();
            this.previousRunningPages = previousRunningPages;
            this.previousDocumentPages = previousDocumentPages;
            emitGeneratedStructure = !isRunningContent && options.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers;
            sectionDefinitions = sections ?? System.Array.Empty<SectionBlock>();
            sectionPageNumbers = resolvedSectionPages ?? new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.Ordinal);
            deferredMaterializations = materializations ?? new System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>>();
            optionsStack.Push(options);
            pageGroupStack.Push(0);
        }

        public LayoutResult Layout(IEnumerable<IPdfBlock> blocks) {
            try {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessBlocks(blocks);
                cancellationToken.ThrowIfCancellationRequested();
                bool forceRunningPage = pages.Count == 0 && currentPageBaseOptions.HasAnyRunningContent;
                if (forceRunningPage) EnsurePage();
                FlushPage(forceRunningPage || pageDirty || HasCurrentPageNonContentObjects());

                var result = new LayoutResult(pageContents) { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
                foreach (var p in pages) result.Pages.Add(p);
                result.HasTableOfContents = encounteredTableOfContents;
                result.SectionDefinitions.AddRange(encounteredSectionDefinitions);
                pageContentsTransferred = true;
                return result;
            } catch {
                if (ownsPageContents) pageContents.Dispose();
                throw;
            }
        }

        public void Dispose() {
            if (ownsPageContents && !pageContentsTransferred) pageContents.Dispose();
        }

        private void StartPage(PdfOptions options) {
            if (isRunningContent) throw new InvalidOperationException("Running PDF content cannot create another page.");
            options.Validate();
            int? effectiveMaximumPages = maximumGeneratedPages is int documentMaximum
                ? options.MaxGeneratedPages is int pageMaximum ? System.Math.Min(documentMaximum, pageMaximum) : documentMaximum
                : options.MaxGeneratedPages;
            if (effectiveMaximumPages is int maximumPages && startedPageCount >= maximumPages)
                throw new InvalidDataException("PDF layout exceeded the configured generated page limit.");
            startedPageCount++;
            currentPageBaseOptions = options;
            currentVisiblePageNumber = ResolveNextVisiblePageNumber(pages.Count,
                !emittedPageGroups.Contains(currentPageGroupId), previousVisiblePageNumber, options);
            currentOpts = options;
            if (options.MirrorMargins && currentVisiblePageNumber % 2 == 0) {
                // Keep one effective frame per source options instance. Deep copies
                // retain document assets and font state, so they must not grow per page.
                if (!mirroredPageOptions.TryGetValue(options, out PdfOptions? mirrored)) {
                    mirrored = options.Clone();
                    mirrored.MarginLeft = options.MarginRight;
                    mirrored.MarginRight = options.MarginLeft;
                    mirroredPageOptions.Add(options, mirrored);
                }
                currentOpts = mirrored;
            }
            width = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            yStart = currentOpts.PageHeight - currentOpts.MarginTop;
            y = yStart;
            currentPage = new LayoutResult.Page { Options = currentOpts, PageGroupId = currentPageGroupId };
            sb.Clear();
            behindTextCanvases.Clear();
            pageDirty = false;
            PrepareAndRenderRunningContents();
            for (int i = 0; i < activeLayers.Count; i++) {
                BeginLayerContent(activeLayers[i]);
            }
        }

        private void EnsurePage() {
            if (currentPage == null) StartPage(currentPageBaseOptions);
        }

        private void PadSectionStart(PdfOptions sectionOptions) {
            PdfPageParity? parity = sectionOptions.PageStartParity;
            if (!parity.HasValue || pages.Count == 0) return;
            int nextNumber = sectionOptions.UseContinuingPageNumberForStartParity ? previousVisiblePageNumber + 1 : pages.Count + 1;
            bool nextPageIsEven = nextNumber % 2 == 0;
            if (nextPageIsEven == (parity == PdfPageParity.Even)) return;
            PdfOptions previous = pages[pages.Count - 1].Options;
            StartPage(new PdfOptions {
                PageWidth = previous.PageWidth,
                PageHeight = previous.PageHeight,
                MarginLeft = previous.MarginLeft,
                MarginRight = previous.MarginRight,
                MarginTop = previous.MarginTop,
                MarginBottom = previous.MarginBottom,
                ShowHeader = false,
                ShowPageNumbers = false
            });
            FlushPage(force: true);
        }

        private bool HasCurrentPageNonContentObjects() =>
            currentPage != null &&
            (currentPage.Images.Count > 0 ||
            currentPage.Annotations.Count > 0 ||
            currentPage.TextAnnotations.Count > 0 ||
            currentPage.FreeTextAnnotations.Count > 0 ||
            currentPage.HighlightAnnotations.Count > 0 ||
            currentPage.FormFields.Count > 0 ||
            currentPage.GraphicsStates.Count > 0 ||
            currentPage.Shadings.Count > 0 ||
            currentPage.NamedDestinations.Count > 0);

        private void FlushPage(bool force = false) {
            if (currentPage == null) return;
            if (!force && !pageDirty && !HasCurrentPageNonContentObjects()) {
                currentPage = null;
                sb.Clear();
                pageDirty = false;
                return;
            }
            for (int i = activeLayers.Count - 1; i >= 0; i--) {
                sb.Append("EMC\n");
            }
            if (behindTextCanvases.Count > 0) {
                var background = new StringBuilder();
                foreach (var canvas in behindTextCanvases.OrderBy(item => item.ZOrder)) background.Append(canvas.Content);
                sb.Insert(0, background.ToString());
                behindTextCanvases.Clear();
            }
            currentPage.Content = pageContents.Store(sb);
            pages.Add(currentPage);
            emittedPageGroups.Add(currentPage.PageGroupId);
            emittedPageGroupCounts.TryGetValue(currentPage.PageGroupId, out int groupCount);
            emittedPageGroupCounts[currentPage.PageGroupId] = groupCount + 1;
            previousVisiblePageNumber = currentVisiblePageNumber;
            currentPage = null;
            // Reuse the buffer across pages (content already captured above) instead of re-growing a new
            // StringBuilder per page.
            sb.Clear();
            pageDirty = false;
        }

        private void NewPage(bool preserveEmptyPage = false) {
            cancellationToken.ThrowIfCancellationRequested();
            if (isRunningContent) throw new InvalidOperationException("Running PDF content cannot create another page.");
            if (activeColumnFlow != null) {
                AdvanceColumnFrame(forcePhysicalPage: false, preserveEmptyPage);
                return;
            }
            PrepareActiveContainerScopesForPageBreak();
            FlushPage(preserveEmptyPage || pageDirty || HasCurrentPageNonContentObjects());
            StartPage(currentPageBaseOptions);
            ResumeActiveContainerScopesOnNewPage();
        }

        private double ResolveTopLevelSpacingBefore(double spacingBefore) {
            return y < GetCurrentFramePageStartY() - 0.001 ? spacingBefore : 0D;
        }

        private static double ResolveColumnSpacingBefore(double spacingBefore, double consumed) {
            return consumed > 0.001 ? spacingBefore : 0D;
        }

        private void BeginLayerContent(PdfLayerDefinition definition) {
            if (currentPage == null) return;
            if (!currentPage.Layers.Contains(definition)) currentPage.Layers.Add(definition);
            sb.Append("/OC /").Append(definition.ResourceName).Append(" BDC\n");
        }

    }
}
