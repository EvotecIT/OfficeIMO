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
        private readonly System.Collections.Generic.Stack<PdfCreateOptions> optionsStack = new System.Collections.Generic.Stack<PdfCreateOptions>();
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
        private readonly HashSet<string> tableOfContentsDestinations = new(StringComparer.Ordinal);
        private PdfCreateOptions currentOpts;
        private PdfCreateOptions currentPageBaseOptions;
        private readonly Dictionary<PdfCreateOptions, PdfCreateOptions> mirroredPageOptions = new();
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
        private bool pageHasContent;
        private bool pageHasBodyContent;
        private int runningNonContentObjectCount;
        private int? currentRunningSectionId;
        private PdfCreateOptions? nextRunningPageOptions;
        // Body emission records logical ownership at the same point as page
        // dirtiness, after any paragraph/table pagination has taken place.
        private bool pageDirty {
            get => pageHasContent;
            set {
                pageHasContent = value;
                if (value && !isRunningContent) pageHasBodyContent = true;
                if (value) RecordRunningContentSection();
            }
        }

        // Blank body flow can occupy a source section's page without painting it.
        // Track that ownership separately from whether the page must be emitted.
        private void RecordRunningContentSection() {
            if (!isRunningContent && currentPage != null && currentRunningSectionId is int sectionId) {
                currentPage.RunningContentSectionId ??= sectionId;
                (currentPage.RunningContentSections ??= new HashSet<int>()).Add(sectionId);
            }
        }
        private bool usedBold;
        private bool usedItalic;
        private bool usedBoldItalic;
        private int _canvasClipDepth;
        private bool _suppressCanvasAccessibilityWrappers;
        private bool _suppressCanvasStructureRegistration;
        private bool _suppressCanvasActualTextChildren;
        // Drawings in one pass measure with the same faces; share their parsed canvases.
        private readonly DrawingTextMeasurementCanvases drawingTextCanvases = new DrawingTextMeasurementCanvases();
        private PageStructElement? _canvasStructureParentElement;
        private readonly System.Collections.Generic.List<FlowSemanticScope> flowSemanticScopes = new System.Collections.Generic.List<FlowSemanticScope>();
        private bool stopDocumentFlow;
        private readonly System.Collections.Generic.HashSet<PdfLayoutPositionCapture> initializedPositionCaptures = new System.Collections.Generic.HashSet<PdfLayoutPositionCapture>();
        private readonly System.Collections.Generic.List<PdfLayerDefinition> activeLayers = new System.Collections.Generic.List<PdfLayerDefinition>();
        private readonly System.Collections.Generic.List<ContainerRenderScope> activeContainerScopes = new System.Collections.Generic.List<ContainerRenderScope>();
        private readonly System.Collections.Generic.Dictionary<(string Key, string Type, PageStructElement? Parent, string Scope, int Columns, int Rows, string AlternativeText, string? Language, PdfListNumbering? Numbering), PageStructElement> canvasStructureElements = new System.Collections.Generic.Dictionary<(string, string, PageStructElement?, string, int, int, string, string?, PdfListNumbering?), PageStructElement>();

        public LayoutContext(
            PdfCreateOptions options,
            System.Collections.Generic.IReadOnlyList<SectionBlock>? sections = null,
            System.Collections.Generic.IReadOnlyDictionary<string, int>? resolvedSectionPages = null,
            System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>>? materializations = null,
            PdfPageContentStore? sharedPageContents = null,
            bool isRunningContent = false,
            RunningContentImageAssets? runningAssets = null,
            IReadOnlyList<PageNumberInfo>? previousRunningPages = null,
            int previousDocumentPages = 1,
            DrawingPatternAnnotationBudget? drawingPatternAnnotations = null,
            PictureAnchorState? pictureAnchors = null, SvgFormRegistry? svgForms = null,
            System.Threading.CancellationToken cancellationToken = default) {
            this.pictureAnchors = pictureAnchors ?? new PictureAnchorState();
            this.svgForms = svgForms ?? new SvgFormRegistry();
            this.drawingPatternAnnotations = drawingPatternAnnotations ?? new DrawingPatternAnnotationBudget();
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
            // Every text path of this layout, running content included, writes to the current page,
            // so color glyphs take their images and opacity states from that page. Restored even when layout fails.
            IColorGlyphPage? previousColorGlyphPage = ContentStreamBuilder.ExchangeColorGlyphPage(new PageColorGlyphs(() => currentPage));
            try {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessBlocks(blocks);
                cancellationToken.ThrowIfCancellationRequested();
                bool forceRunningPage = pages.Count == 0 && currentPageBaseOptions.HasAnyRunningContent;
                if (forceRunningPage) EnsurePage();
                FlushPage(forceRunningPage || pageDirty || HasCurrentPageNonContentObjects());
                FinishFootnoteContinuations();

                var result = new LayoutResult(pageContents) { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic, SvgForms = svgForms };
                foreach (var p in pages) result.Pages.Add(p);
                result.HasTableOfContents = encounteredTableOfContents;
                result.SectionDefinitions.AddRange(encounteredSectionDefinitions);
                pageContentsTransferred = true;
                return result;
            } catch {
                if (ownsPageContents) pageContents.Dispose();
                throw;
            } finally {
                ContentStreamBuilder.ExchangeColorGlyphPage(previousColorGlyphPage);
            }
        }

        public void Dispose() {
            if (ownsPageContents && !pageContentsTransferred) pageContents.Dispose();
        }

        private bool ContinuesCurrentPage(PageBlock pageBlock) =>
            pageBlock.Options.ContinuesCurrentPage && currentPage != null && !isRunningContent && activeColumnFlow == null &&
            activeContainerScopes.Count == 0 &&
            System.Math.Abs(currentPage.Options.PageWidth.ToPoints() - pageBlock.Options.PageWidth.ToPoints()) < .01D &&
            System.Math.Abs(currentPage.Options.PageHeight.ToPoints() - pageBlock.Options.PageHeight.ToPoints()) < .01D;

        // The new group takes the page's remaining height with its own side margins; the
        // next page starts with all of its own setup.
        private void ContinuePageGroupOnCurrentPage(PdfCreateOptions options) {
            options.Validate();
            currentPageBaseOptions = options;
            bool mirrored = options.MirrorMargins && currentVisiblePageNumber % 2 == 0;
            PdfCreateOptions frame = options.Clone();
            frame.MarginLeft = mirrored ? options.MarginRight : options.MarginLeft;
            frame.MarginRight = mirrored ? options.MarginLeft : options.MarginRight;
            frame.MarginTop = currentPage!.Options.MarginTop;
            frame.MarginBottom = currentPage.Options.MarginBottom;
            currentOpts = frame;
            continuedPageAnchorOptions = frame;
            width = frame.PageWidth.ToPoints() - frame.MarginLeft.ToPoints() - frame.MarginRight.ToPoints();
            (currentPage.ContinuedPageGroups ??= new()).Add((currentPageGroupId, frame));
            if (currentRunningSectionId is int sectionId) (currentPage.RunningContentSections ??= new()).Add(sectionId);
        }

        private PdfCreateOptions? continuedPageAnchorOptions;

        private void StartPage(PdfCreateOptions options) {
            if (isRunningContent) throw new InvalidOperationException("Running PDF content cannot create another page.");
            continuedPageAnchorOptions = null;
            if (nextRunningPageOptions != null) {
                options = nextRunningPageOptions;
                nextRunningPageOptions = null;
            }
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
            currentOpts = GetEffectivePageOptions(options, currentVisiblePageNumber);
            width = currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints();
            yStart = currentOpts.PageHeight.ToPoints() - currentOpts.MarginTop.ToPoints();
            y = yStart;
            currentPage = new LayoutResult.Page {
                Options = currentOpts, PageGroupId = currentPageGroupId,
                OperationCancellationToken = cancellationToken,
                // The section at the page start owns its running fields even
                // when it is empty and a continuous successor supplies the body.
                RunningContentSectionId = currentRunningSectionId,
                RunningContentSections = currentRunningSectionId is int sectionId ? new HashSet<int> { sectionId } : null
            };
            sb.Clear();
            behindTextCanvases.Clear();
            inFrontOfTextCanvases.Clear();
            pageDirty = false;
            pageHasBodyContent = false;
            PrepareAndRenderRunningContents();
            runningNonContentObjectCount = CountCurrentPageNonContentObjects();
            StartPageFootnotes();
            for (int i = 0; i < activeLayers.Count; i++) {
                BeginLayerContent(activeLayers[i]);
            }
        }

        private PdfCreateOptions GetEffectivePageOptions(PdfCreateOptions options, int visiblePageNumber) {
            if (!options.MirrorMargins || visiblePageNumber % 2 != 0) return options;
            // Keep one effective frame per source options instance. Deep copies
            // retain document assets and font state, so they must not grow per page.
            if (!mirroredPageOptions.TryGetValue(options, out PdfCreateOptions? mirrored)) {
                mirrored = options.Clone();
                mirrored.MarginLeft = options.MarginRight;
                mirrored.MarginRight = options.MarginLeft;
                mirroredPageOptions.Add(options, mirrored);
            }
            return mirrored;
        }

        private void EnsurePage() {
            if (currentPage == null) StartPage(currentPageBaseOptions);
        }

        private void PadSectionStart(PdfCreateOptions sectionOptions) {
            PdfPageParity? parity = sectionOptions.PageStartParity;
            if (!parity.HasValue || pages.Count == 0) return;
            int nextNumber = sectionOptions.UseContinuingPageNumberForStartParity ? previousVisiblePageNumber + 1 : pages.Count + 1;
            bool nextPageIsEven = nextNumber % 2 == 0;
            if (nextPageIsEven == (parity == PdfPageParity.Even)) return;
            PdfCreateOptions previous = pages[pages.Count - 1].Options;
            StartPage(new PdfCreateOptions {
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

        private int CountCurrentPageNonContentObjects() => currentPage == null ? 0 :
            currentPage.Images.Count + currentPage.Annotations.Count + currentPage.TextAnnotations.Count +
            currentPage.FreeTextAnnotations.Count + currentPage.HighlightAnnotations.Count + currentPage.FormFields.Count +
            currentPage.GraphicsStates.Count + currentPage.Shadings.Count + currentPage.NamedDestinations.Count;

        private bool HasCurrentPageNonContentObjects() => CountCurrentPageNonContentObjects() > 0;

        private bool HasCurrentPageBodyContent() => pageHasBodyContent || HasPageFootnotes || y < yStart - .001D ||
            CountCurrentPageNonContentObjects() > runningNonContentObjectCount;

        private void FlushPage(bool force = false) {
            if (currentPage == null) return;
            FlushPageFootnotes();
            if (!force && !pageDirty && !HasCurrentPageNonContentObjects()) {
                currentPage = null;
                sb.Clear();
                pageDirty = false;
                return;
            }
            for (int i = activeLayers.Count - 1; i >= 0; i--) {
                sb.Append("EMC\n");
            }
            ApplyDeferredTextLayers();
            ReportFreeTextAnnotationDiagnostics();
            BindSvgPageImages();
            foreach (PictureParagraphPlanKey key in pictureAnchors.Paragraphs.Keys.Where(key => key.Page == PicturePage).ToArray())
                pictureAnchors.Paragraphs.Remove(key);
            // Text set in a continued group's frame uses the page's font resources.
            if (currentPage.ContinuedPageGroups != null)
                foreach (var continued in currentPage.ContinuedPageGroups) currentPage.Options.MergeFontProgramUsageFrom(continued.Options);
            currentPage.Content = pageContents.Store(sb);
            pages.Add(currentPage);
            emittedPageGroups.Add(currentPage.PageGroupId);
            emittedPageGroupCounts.TryGetValue(currentPage.PageGroupId, out int groupCount);
            emittedPageGroupCounts[currentPage.PageGroupId] = groupCount + 1;
            previousVisiblePageNumber = currentVisiblePageNumber;
            if (currentPage.ContinuedPageGroups != null) {
                foreach (var continued in currentPage.ContinuedPageGroups) {
                    emittedPageGroups.Add(continued.GroupId);
                    emittedPageGroupCounts.TryGetValue(continued.GroupId, out int continuedCount);
                    emittedPageGroupCounts[continued.GroupId] = continuedCount + 1;
                    if (continued.Options.HasExplicitPageNumberStart) previousVisiblePageNumber = continued.Options.PageNumberStart;
                }
            }
            currentPage = null;
            // Reuse the buffer across pages (content already captured above) instead of re-growing a new
            // StringBuilder per page.
            sb.Clear();
            pageDirty = false;
        }

        private void NewPage(bool preserveEmptyPage = false, bool forcePhysicalPage = false, double? firstBodyHeight = null,
            bool firstBodyAllowsFooterOverlap = false, bool firstBodyAllowsExactOverflow = false) {
            cancellationToken.ThrowIfCancellationRequested();
            if (isRunningContent) throw new InvalidOperationException("Running PDF content cannot create another page.");
            double? previousBodyHeight = runningFirstBodyHeight;
            bool previousFooterOverlap = runningFirstBodyAllowsFooterOverlap;
            bool previousExactOverflow = runningFirstBodyAllowsExactOverflow;
            runningFirstBodyHeight = firstBodyHeight;
            runningFirstBodyAllowsFooterOverlap = firstBodyAllowsFooterOverlap;
            runningFirstBodyAllowsExactOverflow = firstBodyAllowsExactOverflow;
            try {
                if (activeColumnFlow != null) {
                    AdvanceColumnFrame(forcePhysicalPage, preserveEmptyPage);
                    return;
                }
                PrepareActiveContainerScopesForPageBreak();
                if (preserveEmptyPage) RecordRunningContentSection();
                FlushPage(preserveEmptyPage || pageDirty || HasCurrentPageNonContentObjects());
                StartPage(currentPageBaseOptions);
                ResumeActiveContainerScopesOnNewPage();
            } finally {
                runningFirstBodyHeight = previousBodyHeight;
                runningFirstBodyAllowsFooterOverlap = previousFooterOverlap;
                runningFirstBodyAllowsExactOverflow = previousExactOverflow;
            }
        }

        private double ResolveTopLevelSpacingBefore(double spacingBefore) {
            return y < GetCurrentFramePageStartY() - 0.001 ? spacingBefore : 0D;
        }

        private double ResolveContainerSpacingBefore(PdfPanelStyle style) =>
            style.ApplySpacingBeforeAtTop ? style.SpacingBefore.ToPoints() : ResolveTopLevelSpacingBefore(style.SpacingBefore.ToPoints());

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
