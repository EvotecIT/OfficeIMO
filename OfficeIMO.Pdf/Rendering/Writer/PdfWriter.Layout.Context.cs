using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext : IDisposable {
        private StringBuilder sb = new StringBuilder();
        private readonly PdfPageContentStore pageContents;
        private bool pageContentsTransferred;
        private readonly System.Collections.Generic.List<LayoutResult.Page> pages = new System.Collections.Generic.List<LayoutResult.Page>();
        private readonly System.Collections.Generic.Stack<PdfOptions> optionsStack = new System.Collections.Generic.Stack<PdfOptions>();
        private readonly System.Collections.Generic.Stack<int> pageGroupStack = new System.Collections.Generic.Stack<int>();
        private readonly System.Collections.Generic.HashSet<string> emittedTableCellNamedDestinations = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
        private readonly bool emitGeneratedStructure;
        private readonly System.Collections.Generic.IReadOnlyList<SectionBlock> sectionDefinitions;
        private readonly System.Collections.Generic.IReadOnlyDictionary<string, int> sectionPageNumbers;
        private readonly System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>> deferredMaterializations;
        private readonly System.Collections.Generic.List<SectionBlock> encounteredSectionDefinitions = new System.Collections.Generic.List<SectionBlock>();
        private readonly System.Collections.Generic.Dictionary<System.Collections.Generic.List<ColItem>, double[]> rowColumnKeepChainHeights = new System.Collections.Generic.Dictionary<System.Collections.Generic.List<ColItem>, double[]>();
        private bool encounteredTableOfContents;
        private PdfOptions currentOpts;
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
        private int? _canvasStructureParentElementIndex;
        private bool stopDocumentFlow;
        private readonly System.Collections.Generic.HashSet<PdfLayoutPositionCapture> initializedPositionCaptures = new System.Collections.Generic.HashSet<PdfLayoutPositionCapture>();
        private readonly System.Collections.Generic.List<PdfLayerDefinition> activeLayers = new System.Collections.Generic.List<PdfLayerDefinition>();

        public LayoutContext(
            PdfOptions options,
            System.Collections.Generic.IReadOnlyList<SectionBlock>? sections = null,
            System.Collections.Generic.IReadOnlyDictionary<string, int>? resolvedSectionPages = null,
            System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>>? materializations = null) {
            currentOpts = options;
            pageContents = new PdfPageContentStore(options.PageContentMemoryLimitBytes);
            emitGeneratedStructure = options.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers;
            sectionDefinitions = sections ?? System.Array.Empty<SectionBlock>();
            sectionPageNumbers = resolvedSectionPages ?? new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.Ordinal);
            deferredMaterializations = materializations ?? new System.Collections.Generic.Dictionary<FlowMaterializationKey, System.Collections.Generic.IReadOnlyList<IPdfBlock>>();
            optionsStack.Push(options);
            pageGroupStack.Push(0);
        }

        public LayoutResult Layout(IEnumerable<IPdfBlock> blocks) {
            try {
                ProcessBlocks(blocks);
                FlushPage(pageDirty || HasCurrentPageNonContentObjects());

                var result = new LayoutResult(pageContents) { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
                foreach (var p in pages) result.Pages.Add(p);
                result.HasTableOfContents = encounteredTableOfContents;
                result.SectionDefinitions.AddRange(encounteredSectionDefinitions);
                pageContentsTransferred = true;
                return result;
            } catch {
                pageContents.Dispose();
                throw;
            }
        }

        public void Dispose() {
            if (!pageContentsTransferred) pageContents.Dispose();
        }

        private void StartPage(PdfOptions options) {
            options.Validate();
            currentOpts = options;
            width = options.PageWidth - options.MarginLeft - options.MarginRight;
            yStart = options.PageHeight - options.MarginTop;
            y = yStart;
            currentPage = new LayoutResult.Page { Options = options, PageGroupId = currentPageGroupId };
            sb.Clear();
            pageDirty = false;
            for (int i = 0; i < activeLayers.Count; i++) {
                BeginLayerContent(activeLayers[i]);
            }
        }

        private void EnsurePage() {
            if (currentPage == null) StartPage(currentOpts);
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
            currentPage.Content = pageContents.Store(sb.ToString());
            pages.Add(currentPage);
            currentPage = null;
            sb = new StringBuilder();
            pageDirty = false;
        }

        private void NewPage() {
            FlushPage(pageDirty || HasCurrentPageNonContentObjects());
            StartPage(currentOpts);
        }

        private double ResolveTopLevelSpacingBefore(double spacingBefore) {
            return y < yStart - 0.001 ? spacingBefore : 0D;
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
