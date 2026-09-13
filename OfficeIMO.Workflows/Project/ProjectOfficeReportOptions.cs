using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>Controls chart pages and editable data in Word and PowerPoint reports.</summary>
public sealed class ProjectOfficeReportOptions {
    /// <summary>Include rendered charts for visual views. Table views already contain editable tables and are not duplicated as pictures.</summary>
    public bool IncludeCharts { get; set; } = true;
    /// <summary>Include editable report values and supplemental tables. A Table view always retains its primary editable content.</summary>
    public bool IncludeDataTables { get; set; } = true;
    /// <summary>Chart image density and font configuration. Defaults to 300 DPI.</summary>
    public ProjectImageExportOptions Images { get; set; } = new();
}
