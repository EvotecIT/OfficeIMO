using OfficeIMO.Data;

namespace OfficeIMO.Excel;

/// <summary>Controls native Excel table creation and optional bounded nested-object projection.</summary>
public sealed class ExcelTableOptions : OfficeTableOptions {
    /// <summary>Gets or sets flattening, column selection, formatting and materialization limits.</summary>
    public ObjectFlattenerOptions? Projection { get; set; }
    /// <summary>Gets or sets whether eligible table columns are automatically fitted. Defaults to true.</summary>
    public bool AutoFit { get; set; } = true;
    // Mechanical callers of positional headerless insertion retain native Column1, Column2 names.
    internal bool UseDefaultHeaderlessColumnNames { get; set; }
    /// <summary>Gets or sets the native table style.</summary>
    public ExcelTableStyle Style { get; set; } = ExcelTableStyle.TableStyleMedium2;
    /// <summary>Gets or sets whether the table has filter buttons.</summary>
    public bool IncludeAutoFilter { get; set; } = true;
    /// <summary>Gets or sets the execution policy for DataTable cell writes. Null uses the worksheet policy.</summary>
    public ExcelExecutionMode? ExecutionMode { get; set; }
    /// <summary>Gets or sets cancellation for DataTable insertion.</summary>
    public System.Threading.CancellationToken CancellationToken { get; set; }
}
