using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a styled rich-text run to the enclosing cell or comment.</summary>
/// <example>
/// <summary>Add a bold run to a cell</summary>
/// <code>New-OfficeExcel -Path .\Status.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelCell A1 {
///             ExcelText 'Owner: ' -Bold
///             ExcelText 'Infrastructure team'
///         }
///     }
/// }</code>
/// <para>Builds a cell value from a bold label and plain text.</para>
/// </example>
/// <example>
/// <summary>Highlight a status and add a footnote marker</summary>
/// <code>Open-OfficeExcel -Path .\Reports\Migration.xlsx {
///     ExcelSheet 'Overview' {
///         ExcelCell B3 {
///             ExcelText 'Cutover risk: '
///             ExcelText 'High' -Bold -Color Danger -Size 12pt
///             ExcelText '1' -VerticalAlignment Superscript
///         }
///         ExcelCell B20 '1 Vendor patch not yet certified for production.'
///     }
/// }</code>
/// <para>Colours the risk word by its theme role and adds a superscript reference to the notes below.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelText", DefaultParameterSetName = "Default")]
[Alias("ExcelText")]
public sealed class AddOfficeExcelTextCommand : PSCmdlet {
    /// <summary>Comment text.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public string Text { get; set; } = null!;

    /// <summary>Applies bold text. An explicit false value clears bold; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Bold { get; set; }

    /// <summary>Applies italic text. An explicit false value clears italic; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Italic { get; set; }

    /// <summary>Applies strikethrough; explicit false clears it.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Strikethrough { get; set; }

    /// <summary>Apply or clear underline styling across the target row span.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelUnderlineStyle? Underline { get; set; }

    /// <summary>Text or rule colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }

    /// <summary>Font size, interpreted as points for numbers; explicit physical units are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Normal, superscript or subscript positioning for the text run.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelVerticalTextAlignment? VerticalAlignment { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var run = new ExcelRichTextRun(Text) { Color = Color, FontFamily = Font, FontSize = Size, VerticalTextAlignment = VerticalAlignment };
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Bold))) run.Bold = Bold;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Italic))) run.Italic = Italic;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Strikethrough))) run.Strikethrough = Strikethrough;
        if (Underline.HasValue) run.UnderlineStyle = Underline.Value;
        WriteObject(run);
    }

}
