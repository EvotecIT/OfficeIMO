using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

/// <summary>Contains an Excel conversion and the Numbers source evidence that produced it.</summary>
public sealed class NumbersToExcelResult : IDisposable {
    internal NumbersToExcelResult(ExcelDocument value, IWorkSourceDocument source,
        IWorkNumbersProjection projection, IWorkConversionReport report) {
        Value = value;
        Source = source;
        Projection = projection;
        Report = report;
    }

    /// <summary>Gets the converted editable OfficeIMO Excel workbook.</summary>
    public ExcelDocument Value { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Numbers source projection.</summary>
    public IWorkNumbersProjection Projection { get; }
    /// <summary>Gets the loss-aware conversion report.</summary>
    public IWorkConversionReport Report { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => Report.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened, omitted, or reported as errors.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the converted workbook.</summary>
    public ExcelDocument RequireValue() => Value;

    /// <summary>Returns the converted workbook or throws when the conversion was lossy.</summary>
    public ExcelDocument RequireNoLoss() {
        try {
            Report.RequireNoLoss();
            return Value;
        } catch {
            Value.Dispose();
            throw;
        }
    }

    /// <inheritdoc />
    public void Dispose() => Value.Dispose();
}
