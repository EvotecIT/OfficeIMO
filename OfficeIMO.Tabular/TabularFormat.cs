namespace OfficeIMO.Tabular;

/// <summary>
/// Identifies the physical format behind a tabular reader.
/// </summary>
public enum TabularFormat {
    /// <summary>Detect the format from a file name.</summary>
    Auto,

    /// <summary>Delimited text such as CSV or TSV.</summary>
    DelimitedText,

    /// <summary>Open XML Excel workbooks such as XLSX and XLSM.</summary>
    ExcelOpenXml,

    /// <summary>Binary Excel workbooks stored as XLSB packages.</summary>
    ExcelBinary
}
