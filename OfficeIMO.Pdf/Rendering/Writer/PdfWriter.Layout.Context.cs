using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly StringBuilder sb = new StringBuilder();
        private readonly System.Collections.Generic.List<LayoutResult.Page> pages = new System.Collections.Generic.List<LayoutResult.Page>();
        private readonly System.Collections.Generic.Stack<PdfOptions> optionsStack = new System.Collections.Generic.Stack<PdfOptions>();
        private readonly System.Collections.Generic.Stack<int> pageGroupStack = new System.Collections.Generic.Stack<int>();
        private readonly System.Collections.Generic.HashSet<string> emittedTableCellNamedDestinations = new System.Collections.Generic.HashSet<string>(System.StringComparer.Ordinal);
        private readonly bool emitGeneratedStructure;
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

        public LayoutContext(PdfOptions options) {
            currentOpts = options;
            emitGeneratedStructure = options.TaggedStructureMode == PdfTaggedStructureMode.CatalogMarkers;
            optionsStack.Push(options);
            pageGroupStack.Push(0);
        }

        public LayoutResult Layout(IEnumerable<IPdfBlock> blocks) {
            ProcessBlocks(blocks);
            FlushPage(pageDirty || HasCurrentPageNonContentObjects());

            var result = new LayoutResult { UsedBold = usedBold, UsedItalic = usedItalic, UsedBoldItalic = usedBoldItalic };
            foreach (var p in pages) result.Pages.Add(p);
            return result;
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
            currentPage.Content = sb.ToString();
            pages.Add(currentPage);
            currentPage = null;
            sb.Clear();
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

    }
}
