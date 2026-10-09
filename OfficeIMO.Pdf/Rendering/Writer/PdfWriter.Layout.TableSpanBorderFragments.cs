namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Retains source border sides while closing a merged cell at a frame boundary.</summary>
    private static PdfCellBorder PrepareTableSpanFragmentBorder(PdfCellBorder source, TableSpanCellFlow span,
        bool continuation, bool continues) {
        PdfCellBorder border = source.Clone();
        PdfCellBorderSide? top = ResolveCellBorderSide(source.TopBorderSnapshot, source);
        PdfCellBorderSide? bottom = ResolveCellBorderSide(source.BottomBorderSnapshot, source);
        if (continuation && source.Top && IsRenderableCellBorderSide(top)) {
            border.Top = true;
            border.TopBorder = top;
            border.HiddenTopColumnSegments = null;
        }
        if (continues && source.Bottom && IsRenderableCellBorderSide(bottom)) {
            border.Bottom = true;
            border.BottomBorder = bottom;
            border.HiddenBottomColumnSegments = null;
        }
        int offset = span.FirstFragmentRow - span.Row;
        int count = span.LastFragmentRow - span.FirstFragmentRow + 1;
        HashSet<int>? Slice(HashSet<int>? hidden) => hidden == null ? null :
            new HashSet<int>(hidden.Where(row => row >= offset && row < offset + count).Select(row => row - offset));
        border.HiddenLeftRowSegments = Slice(source.HiddenLeftRowSegments);
        border.HiddenRightRowSegments = Slice(source.HiddenRightRowSegments);
        return border;
    }
}
