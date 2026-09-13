using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>A contiguous set of table rows and its measured height, including the repeated header.</summary>
public sealed class OfficeTablePage {
    internal OfficeTablePage(int offset, int count, double height) { RowOffset = offset; RowCount = count; Height = height; }
    /// <summary>Zero-based first source row.</summary>
    public int RowOffset { get; }
    /// <summary>Number of source rows on this page.</summary>
    public int RowCount { get; }
    /// <summary>Total row height plus the repeated header, in caller-defined units.</summary>
    public double Height { get; }
}

/// <summary>Shared measured-row pagination for drawing and native document tables.</summary>
public static class OfficeTablePagination {
    /// <summary>
    /// Packs consecutive measured rows without splitting or dropping them. Every page includes headerHeight.
    /// Invalid dimensions, an unfit row, or an exceeded page budget fail before a partial result is returned.
    /// An empty table produces one header-only page.
    /// </summary>
    public static IReadOnlyList<OfficeTablePage> Paginate(IReadOnlyList<double> rowHeights, double availableHeight,
        double headerHeight, int maximumPages = 1000, CancellationToken cancellationToken = default) {
        if (rowHeights == null) throw new ArgumentNullException(nameof(rowHeights));
        if (!Finite(availableHeight) || availableHeight <= 0 || !Finite(headerHeight) || headerHeight < 0 || headerHeight > availableHeight)
            throw new ArgumentOutOfRangeException(nameof(availableHeight), "The repeated header must fit inside a positive page height.");
        if (maximumPages < 1) throw new ArgumentOutOfRangeException(nameof(maximumPages));
        var pages = new List<OfficeTablePage>();
        int offset = 0, count = 0; double height = headerHeight;
        for (int i = 0; i < rowHeights.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            double row = rowHeights[i];
            if (!Finite(row) || row <= 0) throw new ArgumentOutOfRangeException(nameof(rowHeights), "Row heights must be finite and positive.");
            if (row > availableHeight - headerHeight) throw new InvalidOperationException("A table row cannot fit with its repeated header. Increase the page height or column width.");
            if (row > availableHeight - height && count > 0) {
                pages.Add(new OfficeTablePage(offset, count, height));
                if (pages.Count >= maximumPages) throw new InvalidOperationException("Table pagination exceeds its page limit.");
                offset = i; count = 0; height = headerHeight;
            }
            count++; height += row;
        }
        cancellationToken.ThrowIfCancellationRequested();
        pages.Add(new OfficeTablePage(offset, count, height));
        return pages.AsReadOnly();
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
