using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Security;
using System.Threading.Tasks;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Creates Excel-to-HTML settings for Convert-OfficeDocument -To Html.</summary>
/// <example>
/// <summary>Convert a workbook to HTML</summary>
/// <code>Convert-OfficeDocument -Path .\Reports\Q3-Sales.xlsx -To Html -OutputPath .\Site\Q3-Sales.html -Options (New-OfficeExcelHtmlOptions -Title 'Q3 sales')</code>
/// <para>Converts a workbook to a web page with a title.</para>
/// </example>
/// <example>
/// <summary>Publish a visual review page</summary>
/// <code>$html = New-OfficeExcelHtmlOptions -Profile VisualReview -Theme Boardroom -Title 'Operations dashboard' -MaxRowsPerSheet 500
/// Convert-OfficeDocument -Path .\Reports\Operations.xlsx -To Html -OutputPath .\Site\Operations.html -Options $html</code>
/// <para>Renders the workbook as a visual review page in the Boardroom theme, capped at 500 rows per sheet for a fast page.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelHtmlOptions", DefaultParameterSetName = "Default")]
public sealed class NewOfficeExcelHtmlOptionsCommand : PSCmdlet {
    /// <summary>HTML conversion profile.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeExcelHtmlProfile Profile { get; set; } = OfficeExcelHtmlProfile.SemanticTables;

    /// <summary>Built-in Office theme, native theme object or a document whose visual theme is read.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeThemeTransformation]
    public OfficeThemeInput? Theme { get; set; }

    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Title { get; set; }

    /// <summary>Maximum number of used rows to emit per worksheet.</summary>
    [Parameter(ParameterSetName = "Default")]
    public int? MaxRowsPerSheet { get; set; }

    /// <summary>Text used for empty cells in semantic table output.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? EmptyCellText { get; set; }

    /// <summary>Do not include OfficeIMO default CSS styles.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoDefaultStyles { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.Excel.Html.ExcelHtmlExportOptions {
            ExportProfile = Profile == OfficeExcelHtmlProfile.VisualReview ? OfficeIMO.Excel.Html.ExcelHtmlExportProfile.VisualReview : OfficeIMO.Excel.Html.ExcelHtmlExportProfile.SemanticTables,
        };
        options.DocumentOutput.Title = Title;
        options.DocumentOutput.IncludeDefaultStyles = Theme != null;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(NoDefaultStyles))) options.DocumentOutput.IncludeDefaultStyles = !NoDefaultStyles.IsPresent;
        if (Theme != null) options.DocumentOutput.DocumentTheme = Theme.Resolve(OfficeTheme.FromDocument);
        if (MaxRowsPerSheet.HasValue) options.MaxRowsPerSheet = MaxRowsPerSheet.Value;
        if (EmptyCellText != null) options.EmptyCellText = EmptyCellText;
        WriteObject(options);
    }

}
