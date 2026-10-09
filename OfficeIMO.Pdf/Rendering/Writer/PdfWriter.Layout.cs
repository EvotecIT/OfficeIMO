using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private const double TableCellClipBleed = 2D;
    private const double TableCellCheckBoxGap = 2D;
    private const double TableCellNoWrapWidth = 1000000D;

    // Helper shapes for column pagination
    private abstract class ColItem {
        public string Kind = string.Empty;
        public double ColumnXOffset;
        public double ColumnWidth;
    }
    private sealed class ColPar : ColItem { public RichParagraphBlock Block = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double XOffset; public double TextWidth; public double FirstLineXOffset; public double FirstLineTextWidth; public ColPar() { Kind = "P"; } }
    private sealed class ColHead : ColItem { public HeadingBlock Block = null!; public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs = null!; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!; public System.Collections.Generic.List<double> Heights = null!; public double Leading; public double Size; public double SpacingBefore; public double SpacingAfter; public bool Bold; public bool ApplySpacingBeforeAtTop; public bool KeepWithNext; public PdfColor? Color; public ColHead() { Kind = "H"; } }
    private sealed class ColRule : ColItem { public HorizontalRuleBlock Block = null!; public ColRule() { Kind = "R"; } }
    private sealed class ColImg : ColItem { public ImageBlock Block = null!; public PdfImageStyle Style = null!; public double Width; public double Height; public ColImg() { Kind = "I"; } }
    private sealed class ColShape : ColItem { public ShapeBlock Block = null!; public ColShape() { Kind = "S"; } }
    private sealed class ColDrawing : ColItem { public DrawingBlock Block = null!; public ColDrawing() { Kind = "D"; } }
    private sealed class ColForm : ColItem { public IPdfBlock Block = null!; public ColForm() { Kind = "FORM"; } }
    private sealed class ColAnnotation : ColItem { public IPdfBlock Block = null!; public ColAnnotation() { Kind = "ANNOTATION"; } }
    private sealed class ColBookmark : ColItem { public BookmarkBlock Block = null!; public ColBookmark() { Kind = "B"; } }
    private sealed class ColSpacer : ColItem { public SpacerBlock Block = null!; public ColSpacer() { Kind = "SPACE"; } }
    private sealed class ColListItem : ColItem {
        public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs = null!;
        public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines = null!;
        public System.Collections.Generic.List<double> Heights = null!;
        public string Marker = string.Empty;
        public PdfStandardFont MarkerFont;
        public PdfNamedFontFace? MarkerNamedFont;
        public double MarkerSize;
        public PdfColor? MarkerColor;
        public double MarkerXOffset;
        public double MarkerWidth;
        public PdfAlign MarkerAlign;
        public double TextXOffset;
        public double TextWidth;
        public PdfAlign TextAlign;
        public PdfColor? Color;
        public double Leading;
        public double Size;
        public double SpacingBefore;
        public double SpacingAfter;
        public string? BookmarkName;
        public int ListGroupId;
        public bool KeepTogether;
        public bool IsFirstInKeepGroup;
        public double KeepGroupHeight;
        public bool KeepWithNext;
        public bool IsFirstInKeepWithNextGroup;
        public int KeepWithNextGroupItemCount;
        public double KeepWithNextGroupHeight;
        public PageStructElement? StructureElement;

        public ColListItem() {
            Kind = "L";
        }
    }
    private sealed class ColTable : ColItem { public TableBlock Block = null!; public PdfTableStyle Style = null!; public int Columns; public double[] ColumnWidths = null!; public TableCellTextLayout[][] RowLines = null!; public int[] RowLineCounts = null!; public double[] RowHeights = null!; public double[] RowLeadings = null!; public double[] RowSizes = null!; public double[] RowRunFontSizeScales = null!; public bool[] RowBold = null!; public double Width; public double Size; public int HeaderRowCount; public int RepeatHeaderRowCount; public int FooterStartRowIndex; public System.Collections.Generic.IReadOnlyList<PdfTextRun>? CaptionRuns; public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>? CaptionLines; public System.Collections.Generic.List<double>? CaptionLineHeights; public double CaptionLeading; public double CaptionHeight; public ColTable() { Kind = "T"; } }
    private sealed class TableColumnLayout { public double[] Widths = null!; public double Width; }
    private sealed class TableCellTextLayout {
        public TableCellTextLayout(System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> lines, System.Collections.Generic.List<double> lineHeights, System.Collections.Generic.List<PdfAlign?>? lineAlignments = null, System.Collections.Generic.List<double>? lineXOffsets = null, System.Collections.Generic.List<double>? lineWidths = null, double topSpacing = 0D, System.Collections.Generic.List<double>? lineBoxHeights = null, System.Collections.Generic.List<TableCellParagraphRange>? paragraphRanges = null) {
            Lines = lines;
            LineHeights = lineHeights;
            LineHeightPrefix = new double[lineHeights.Count + 1];
            for (int index = 0; index < lineHeights.Count; index++)
                LineHeightPrefix[index + 1] = LineHeightPrefix[index] + lineHeights[index];
            LineBoxHeights = lineBoxHeights ?? lineHeights;
            LineAlignments = lineAlignments;
            LineXOffsets = lineXOffsets;
            LineWidths = lineWidths;
            TopSpacing = topSpacing;
            ParagraphRanges = paragraphRanges;
        }

        public System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>> Lines { get; }
        public System.Collections.Generic.List<double> LineHeights { get; }
        public double[] LineHeightPrefix { get; }
        public System.Collections.Generic.List<double> LineBoxHeights { get; }
        public System.Collections.Generic.List<PdfAlign?>? LineAlignments { get; }
        public System.Collections.Generic.List<double>? LineXOffsets { get; }
        public System.Collections.Generic.List<double>? LineWidths { get; }
        public double TopSpacing { get; }
        public System.Collections.Generic.List<TableCellParagraphRange>? ParagraphRanges { get; }
        public int LineCount => System.Math.Max(1, Lines.Count);
    }
    private readonly record struct TableCellParagraphRange(PdfTableCellParagraph Paragraph, int StartLine, int LineCount);
    private readonly struct TableCellLayout {
        public TableCellLayout(int column, int columnSpan, int rowSpan, string text, System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> paragraphs, string? linkUri, string? linkDestinationName, string? linkContents, string? namedDestinationName, System.Collections.Generic.IReadOnlyList<PdfTableCellCheckBox> checkBoxes, System.Collections.Generic.IReadOnlyList<PdfTableCellFormField> formFields, System.Collections.Generic.IReadOnlyList<PdfTableCellImage> images, bool noWrap, PdfTableCellViewport? viewport, int textRotation) {
            Column = column;
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            Text = text;
            Runs = runs;
            Paragraphs = paragraphs;
            LinkUri = linkUri;
            LinkDestinationName = linkDestinationName;
            LinkContents = linkContents;
            NamedDestinationName = namedDestinationName;
            CheckBoxes = checkBoxes;
            FormFields = formFields;
            Images = images;
            NoWrap = noWrap;
            Viewport = viewport;
            TextRotation = textRotation;
        }

        public int Column { get; }
        public int ColumnSpan { get; }
        public int RowSpan { get; }
        public string Text { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> Paragraphs { get; }
        public string? LinkUri { get; }
        public string? LinkDestinationName { get; }
        public string? LinkContents { get; }
        public string? NamedDestinationName { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellCheckBox> CheckBoxes { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellFormField> FormFields { get; }
        public System.Collections.Generic.IReadOnlyList<PdfTableCellImage> Images { get; }
        public bool NoWrap { get; }
        public PdfTableCellViewport? Viewport { get; }
        public int TextRotation { get; }
        public double? OrientedRowTextHeight { get; init; }
    }

    private static LayoutResult LayoutBlocks(
        IEnumerable<IPdfBlock> blocks,
        PdfOptions opts,
        System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var blockList = blocks as IReadOnlyList<IPdfBlock> ?? blocks.ToList();
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<SectionBlock> sections = Array.Empty<SectionBlock>();
        IReadOnlyDictionary<string, int>? pageNumbers = null;
        var deferredMaterializations = new Dictionary<FlowMaterializationKey, IReadOnlyList<IPdfBlock>>();
        var runningAssets = new RunningContentImageAssets();
        IReadOnlyList<PageNumberInfo>? previousRunningPages = null;
        int previousDocumentPages = 1;
        LayoutResult result = null!;
        for (int pass = 0; pass < 5; pass++) {
            cancellationToken.ThrowIfCancellationRequested();
            result?.Dispose();
            using var context = new LayoutContext(opts, sections, pageNumbers, deferredMaterializations,
                runningAssets: runningAssets, previousRunningPages: previousRunningPages,
                previousDocumentPages: previousDocumentPages, cancellationToken: cancellationToken);
            result = context.Layout(blockList);
            IReadOnlyList<PageNumberInfo> runningPages = BuildPageNumberInfos(result.Pages);
            bool runningStable = RunningContentContextsMatch(result, runningPages);
            previousRunningPages = runningPages;
            previousDocumentPages = result.Pages.Count;
            if (!result.HasTableOfContents && runningStable) {
                ApplySectionReferences(result);
                return result;
            }

            IReadOnlyDictionary<string, int> resolved = BuildSectionPageNumbers(result);
            IReadOnlyList<SectionBlock> resolvedSections = result.SectionDefinitions;
            if (pageNumbers != null && runningStable &&
                SectionPageNumbersEqual(pageNumbers, resolved) &&
                SectionDefinitionsEqual(sections, resolvedSections)) {
                ApplySectionReferences(result);
                return result;
            }

            pageNumbers = resolved;
            sections = resolvedSections;
        }

        bool hasDynamicRunningContent = result?.Pages.Any(page => page.HasDynamicRunningContent) == true;
        result?.Dispose();
        throw new InvalidOperationException(hasDynamicRunningContent
            ? "Running PDF content did not stabilize within five layout passes."
            : "Generated table of contents did not stabilize within five layout passes.");
    }

    private static bool SectionDefinitionsEqual(IReadOnlyList<SectionBlock> left, IReadOnlyList<SectionBlock> right) {
        if (left.Count != right.Count) return false;
        for (int i = 0; i < left.Count; i++) {
            SectionBlock first = left[i];
            SectionBlock second = right[i];
            if (!string.Equals(first.DestinationName, second.DestinationName, StringComparison.Ordinal) ||
                !string.Equals(first.Title, second.Title, StringComparison.Ordinal) ||
                first.Options.Level != second.Options.Level ||
                first.Options.IncludeInTableOfContents != second.Options.IncludeInTableOfContents) {
                return false;
            }
        }

        return true;
    }

    private static Dictionary<string, int> BuildSectionPageNumbers(LayoutResult result) {
        var pageNumbers = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int pageIndex = 0; pageIndex < result.Pages.Count; pageIndex++) {
            foreach (PageSection section in result.Pages[pageIndex].Sections) {
                pageNumbers[section.DestinationName] = pageIndex + 1;
            }
        }

        return pageNumbers;
    }

    private static bool SectionPageNumbersEqual(IReadOnlyDictionary<string, int> left, IReadOnlyDictionary<string, int> right) {
        if (left.Count != right.Count) return false;
        foreach (KeyValuePair<string, int> pair in left) {
            if (!right.TryGetValue(pair.Key, out int page) || page != pair.Value) return false;
        }

        return true;
    }

    private static void ApplySectionReferences(LayoutResult result) {
        for (int pageIndex = 0; pageIndex < result.Pages.Count; pageIndex++) {
            foreach (PageSection section in result.Pages[pageIndex].Sections) {
                section.Reference?.Set(section.DestinationName, section.Title, pageIndex + 1, section.Y);
            }
        }
    }
}
