using System.Threading;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static double ScaleEditableFontSize(double fontSize, double scale) =>
        Math.Min(4000D, Math.Max(1D, fontSize * scale));

    private static void ApplyEditableTableTypography(
        PptCore.PowerPointTable table,
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfLogicalTable sourceTable,
        double scale,
        CancellationToken cancellationToken) {
        int fontSize = (int)Math.Round(Math.Min(
            18D,
            ScaleEditableFontSize(GetMedianTableFontSize(page, sourceTable), scale)));
        fontSize = Math.Max(1, fontSize);
        for (int rowIndex = 0; rowIndex < table.Rows; rowIndex++) {
            for (int columnIndex = 0; columnIndex < table.Columns; columnIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                PptCore.PowerPointTableCell cell = table.GetCell(rowIndex, columnIndex);
                cell.FontSize = fontSize;
                cell.PaddingLeftPoints = 0.5D;
                cell.PaddingRightPoints = 0.5D;
                cell.PaddingTopPoints = 0.25D;
                cell.PaddingBottomPoints = 0.25D;
                cell.VerticalAlignment = PptCore.PowerPointTextVerticalAlignment.Top;
                cell.TextAutoFit = PptCore.PowerPointTextAutoFit.Normal;
            }
        }
    }

    private static double GetMedianTableFontSize(
        PdfCore.PdfLogicalPage page,
        PdfCore.PdfLogicalTable table) {
        var sizes = new List<double>();
        for (int blockIndex = 0; blockIndex < page.TextBlocks.Count; blockIndex++) {
            PdfCore.PdfLogicalTextBlock block = page.TextBlocks[blockIndex];
            if (block.FontSize > 0D && IsTextInsideTableBounds(block, table)) {
                sizes.Add(block.FontSize);
            }
        }
        if (sizes.Count == 0) return 10D;
        sizes.Sort();
        int middle = sizes.Count / 2;
        return (sizes.Count & 1) == 0
            ? (sizes[middle - 1] + sizes[middle]) / 2D
            : sizes[middle];
    }

    private static bool IsTextInsideTableBounds(
        PdfCore.PdfLogicalTextBlock block,
        PdfCore.PdfLogicalTable table) {
        if (block.VisualBounds is PdfCore.PdfLogicalVisualBounds blockBounds &&
            table.VisualBounds is PdfCore.PdfLogicalVisualBounds tableBounds) {
            double centerX = (blockBounds.Left + blockBounds.Right) / 2D;
            double centerY = (blockBounds.Top + blockBounds.Bottom) / 2D;
            return centerX >= tableBounds.Left && centerX <= tableBounds.Right &&
                centerY >= tableBounds.Top && centerY <= tableBounds.Bottom;
        }
        if (table.Columns.Count == 0 ||
            block.BaselineY < table.YBottom ||
            block.BaselineY > table.YTop) return false;
        double left = table.Columns.Min(static column => column.From);
        double right = table.Columns.Max(static column => column.To);
        return block.XEnd >= left && block.XStart <= right;
    }
}
