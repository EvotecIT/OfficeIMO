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

/// <summary>Adds a paragraph of text to the report.</summary>
/// <example>
/// <summary>Add explanatory text</summary>
/// <code>New-OfficeExcel -Path .\Reports\Operations.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportParagraph 'Figures cover all production services from 1 to 31 October.'
///     }
/// }</code>
/// <para>Adds a paragraph under the title.</para>
/// </example>
/// <example>
/// <summary>Explain a table before it appears</summary>
/// <code>New-OfficeExcel -Path .\Reports\Service-Review-October.xlsx {
///     ExcelReportSheet 'Review' {
///         ExcelReportHeading 'Availability by site'
///         ExcelReportParagraph 'Availability is measured every minute from two external probes. Planned maintenance announced five days ahead is excluded.'
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Site-Availability.csv)
///     }
/// }</code>
/// <para>Describes how availability is measured before the site table, so readers do not need a separate methods document.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportParagraph", DefaultParameterSetName = "Default")]
[Alias("ExcelReportParagraph")]
public sealed class AddOfficeExcelReportParagraphCommand : PSCmdlet {
    /// <summary>Comment text.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Text { get; set; } = null!;

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        context.RequireComposer().Paragraph(Text);
        if (PassThru) WriteObject(context.RequireSheet());
    }

}
