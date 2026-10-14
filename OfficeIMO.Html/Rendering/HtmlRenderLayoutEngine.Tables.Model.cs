using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private sealed class TableCaptionLayout {
        internal TableCaptionLayout(string side, double height, IReadOnlyList<HtmlRenderVisual> visuals) {
            Side = side;
            Height = height;
            Visuals = visuals;
        }

        internal string Side { get; }
        internal double Height { get; }
        internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    }

    private sealed class TableRowLayout {
        internal TableRowLayout(
            IElement element,
            HtmlRenderBoxStyle style,
            IElement? groupElement,
            HtmlRenderBoxStyle? groupStyle,
            IReadOnlyList<TableCellLayout> cells,
            double height,
            bool isHeader,
            bool isFooter, bool anonymous, IElement? continuationElement, string? structureKey) {
            Element = element;
            Style = style;
            GroupElement = groupElement;
            GroupStyle = groupStyle;
            Cells = cells;
            Height = height;
            IsHeader = isHeader;
            IsFooter = isFooter;
            Anonymous = anonymous;
            ContinuationElement = continuationElement;
            StructureKey = structureKey;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal IElement? GroupElement { get; }
        internal HtmlRenderBoxStyle? GroupStyle { get; }
        internal IReadOnlyList<TableCellLayout> Cells { get; }
        internal double Height { get; set; }
        internal double Baseline { get; set; }
        internal bool IsHeader { get; }
        internal bool IsFooter { get; }
        internal bool Anonymous { get; }
        internal IElement? ContinuationElement { get; }
        internal string? StructureKey { get; }
    }

    private sealed class TableCellLayout {
        internal TableCellLayout(IElement element, HtmlRenderBoxStyle style, HtmlInlineLayout inline, int column, int span, int rowSpan, double width, double minimumHeight, bool anonymous, string? structureKey, TableCellPercentageContent? percentageContent) {
            Element = element;
            Style = style;
            Inline = inline;
            Column = column;
            Span = span;
            RowSpan = rowSpan;
            Width = width;
            MinimumHeight = minimumHeight;
            Anonymous = anonymous;
            StructureKey = structureKey;
            PercentageContent = percentageContent;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal HtmlInlineLayout Inline { get; set; }
        internal TableCellPercentageContent? PercentageContent { get; }
        internal int Column { get; }
        internal int Span { get; }
        internal int RowSpan { get; }
        internal double Width { get; }
        internal double MinimumHeight { get; }
        internal bool Anonymous { get; }
        internal string? StructureKey { get; }
        internal double ContentOffsetY { get; set; }
    }
}
