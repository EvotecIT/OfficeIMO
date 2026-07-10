namespace OfficeIMO.Excel.OpenDocument;

/// <summary>Controls bounded content transfer by the Excel/OpenDocument adapter.</summary>
public sealed class ExcelOpenDocumentConversionOptions {
    /// <summary>Copy the basic font, fill, and number-format subset exposed by both typed models.</summary>
    public bool IncludeBasicStyles { get; set; } = true;
    /// <summary>Maximum number of non-empty repeated ODS cells that may be expanded into XLSX cells.</summary>
    public long MaximumExpandedCells { get; set; } = 1_000_000;
    /// <summary>Maximum XLSX row index produced from an ODS sheet.</summary>
    public int MaximumRows { get; set; } = 1_048_576;
    /// <summary>Maximum XLSX column index produced from an ODS sheet.</summary>
    public int MaximumColumns { get; set; } = 16_384;

    internal void Validate() {
        if (MaximumExpandedCells < 1) throw new ArgumentOutOfRangeException(nameof(MaximumExpandedCells));
        if (MaximumRows < 1 || MaximumRows > 1_048_576) throw new ArgumentOutOfRangeException(nameof(MaximumRows));
        if (MaximumColumns < 1 || MaximumColumns > 16_384) throw new ArgumentOutOfRangeException(nameof(MaximumColumns));
    }
}
