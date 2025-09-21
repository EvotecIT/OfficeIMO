namespace OfficeIMO.Pdf;

using System;
using System.Linq;

/// <summary>Row builder with percentage-based columns.</summary>
public class PdfRowCompose {
    private const double WidthTolerance = 0.0001;
    private readonly PdfDoc _doc;
    private readonly RowBlock _row = new RowBlock();
    private double _allocatedWidth;

    internal PdfRowCompose(PdfDoc doc) { _doc = doc; }

    /// <summary>Adds a column with the given width percentage.</summary>
    public PdfRowCompose Column(double widthPercent, System.Action<PdfRowColumnCompose> build) {
        Guard.NotNull(build, nameof(build));

        ValidateWidth(widthPercent);
        EnsureTotalWithinBounds(widthPercent);

        var col = new RowColumn(widthPercent);
        var cc = new PdfRowColumnCompose(col);
        build(cc);
        _row.Columns.Add(col);
        _allocatedWidth += widthPercent;
        return this;
    }

    internal void Commit() {
        NormalizeColumnWidthsIfNeeded();
        _doc.AddRow(_row);
    }

    private static void ValidateWidth(double widthPercent) {
        if (widthPercent <= 0)
            throw new ArgumentOutOfRangeException(nameof(widthPercent), widthPercent, "Column width must be greater than 0%.");

        if (widthPercent > 100)
            throw new ArgumentOutOfRangeException(nameof(widthPercent), widthPercent, "Column width cannot exceed 100%.");
    }

    private void EnsureTotalWithinBounds(double widthPercent) {
        var prospectiveTotal = _allocatedWidth + widthPercent;
        if (prospectiveTotal > 100 + WidthTolerance)
            throw new InvalidOperationException($"Column widths cannot exceed 100%. Current total {_allocatedWidth:F2}% + {widthPercent:F2}%.");
    }

    private void NormalizeColumnWidthsIfNeeded() {
        if (_row.Columns.Count == 0)
            return;

        var total = _row.Columns.Sum(static c => c.WidthPercent);
        if (total <= 0)
            return;

        if (total > 100 + WidthTolerance)
            throw new InvalidOperationException("Row columns exceed 100% total width after composition.");

        if (total >= 100 - WidthTolerance) {
            _allocatedWidth = total;
            return;
        }

        var scale = 100.0 / total;
        double accumulated = 0;
        for (int i = 0; i < _row.Columns.Count; i++) {
            double newWidth;
            if (i == _row.Columns.Count - 1) {
                newWidth = 100 - accumulated;
            } else {
                newWidth = _row.Columns[i].WidthPercent * scale;
                accumulated += newWidth;
            }

            _row.Columns[i].WidthPercent = newWidth;
        }

        _allocatedWidth = _row.Columns.Sum(static c => c.WidthPercent);
    }
}

