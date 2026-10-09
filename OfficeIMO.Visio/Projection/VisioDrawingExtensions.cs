using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Projects cached Visio page content into reusable point-sized Drawing scenes without editing its source.</summary>
public static class VisioDrawingExtensions {
    /// <summary>Projects every source page in document order, including blank and background pages.</summary>
    /// <remarks>Stencil masters must first be instantiated on a page. ShapeSheet recalculation is not performed.</remarks>
    public static OfficeConversionResult<IReadOnlyList<OfficeDrawing>, VisioDrawingConversionReport> ToDrawings(
        this VisioDocument document, VisioDrawingOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        VisioDrawingOptions operation = (options ?? new VisioDrawingOptions()).Snapshot();
        if (document.Pages.Count == 0) throw new InvalidOperationException("The Visio document has no pages. Instantiate stencil masters on a page before rendering.");
        if (document.Pages.Count > operation.MaximumPages) throw new InvalidDataException("Visio page count exceeds MaximumPages.");
        var report = new VisioDrawingConversionReport(operation.SourceName ?? document.FilePath);
        var context = new VisioDrawingProjection(operation, report, cancellationToken);
        var drawings = new List<OfficeDrawing>(document.Pages.Count);
        for (int index = 0; index < document.Pages.Count; index++) drawings.Add(context.Project(document.Pages[index], index + 1));
        if (operation.RequireNoLoss) report.RequireNoLoss();
        return new OfficeConversionResult<IReadOnlyList<OfficeDrawing>, VisioDrawingConversionReport>(drawings.AsReadOnly(), report);
    }

    /// <summary>Projects one page at its physical dimensions in points and returns source-qualified fidelity evidence.</summary>
    public static OfficeConversionResult<OfficeDrawing, VisioDrawingConversionReport> ToDrawing(
        this VisioPage page, VisioDrawingOptions? options = null, CancellationToken cancellationToken = default) {
        if (page == null) throw new ArgumentNullException(nameof(page));
        cancellationToken.ThrowIfCancellationRequested();
        VisioDrawingOptions operation = (options ?? new VisioDrawingOptions()).Snapshot();
        var report = new VisioDrawingConversionReport(operation.SourceName ?? page.OwnerDocument?.FilePath);
        OfficeDrawing drawing = new VisioDrawingProjection(operation, report, cancellationToken).Project(page, 1);
        if (operation.RequireNoLoss) report.RequireNoLoss();
        return new OfficeConversionResult<OfficeDrawing, VisioDrawingConversionReport>(drawing, report);
    }
}
