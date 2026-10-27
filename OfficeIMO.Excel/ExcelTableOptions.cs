using OfficeIMO.Data;

namespace OfficeIMO.Excel;

/// <summary>Controls native Excel table creation and optional bounded nested-object projection.</summary>
public sealed class ExcelTableOptions : OfficeTableOptions {
    /// <summary>Gets or sets flattening, column selection, formatting and materialization limits.</summary>
    public ObjectFlattenerOptions? Projection { get; set; }
    /// <summary>Gets or sets the native table style.</summary>
    public ExcelTableStyle Style { get; set; } = ExcelTableStyle.TableStyleMedium2;
    /// <summary>Gets or sets whether the table has filter buttons.</summary>
    public bool IncludeAutoFilter { get; set; } = true;
}
