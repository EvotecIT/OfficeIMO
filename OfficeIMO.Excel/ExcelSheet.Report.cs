namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Starts a report builder on this worksheet. Existing cells are retained until report content writes them.</summary>
    public ExcelReportBuilder Report(ExcelReportOptions? options = null) => new ExcelReportBuilder(this, options);
}
