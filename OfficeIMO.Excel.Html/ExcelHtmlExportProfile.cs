namespace OfficeIMO.Excel.Html;

/// <summary>Named Excel-to-HTML output contracts.</summary>
public enum ExcelHtmlExportProfile {
    /// <summary>Workbook and worksheet content as accessible semantic tables.</summary>
    SemanticTables,

    /// <summary>Worksheet content and drawing geometry as positioned visual-review HTML.</summary>
    VisualReview
}
