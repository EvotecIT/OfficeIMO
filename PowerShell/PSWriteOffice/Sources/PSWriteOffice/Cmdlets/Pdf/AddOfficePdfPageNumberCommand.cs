using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Inserts a live page number (optionally 'of total') inside PdfHeader or PdfFooter.</summary>
/// <example><summary>Use the native OfficePdfPageNumber command.</summary><code>New-OfficePdf -Path .\Numbered.pdf {
///     PdfFooter {
///         PdfText 'Page '
///         PdfPageNumber -IncludeTotal
///     }
///     PdfParagraph 'Body text.'
/// }</code></example>
[Cmdlet("Add", "OfficePdfPageNumber", DefaultParameterSetName = "Context")]
[OutputType(typeof(void))]
[Alias("PdfPageNumber") ]
public sealed class AddOfficePdfPageNumberCommand : OfficeCmdlet {
    /// <summary>IncludeTotal for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public SwitchParameter IncludeTotal { get; set; }
    /// <summary>Separator for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public string? Separator { get; set; }
    /// <summary>Bold for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public SwitchParameter Bold { get; set; }
    /// <summary>Italic for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public SwitchParameter Italic { get; set; }
    /// <summary>Strikethrough for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public SwitchParameter Strikethrough { get; set; }
    /// <summary>Shared underline style; bare -Underline means Single.</summary>
    [Parameter(ParameterSetName = "Context")]
    public OfficeTextDecorationStyle? Underline { get; set; }
    /// <summary>Color for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }
    /// <summary>BackgroundColor for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? BackgroundColor { get; set; }
    /// <summary>Font family name or standard PDF font; completes from the theme and registered fonts.</summary>
    [Parameter(ParameterSetName = "Context")]
    public string? Font { get; set; }
    /// <summary>Size for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }
    /// <summary>Baseline for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    public PdfTextBaseline? Baseline { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var running = OfficeDocumentContext.Find<PdfRunningTextContent>() ?? throw new PSInvalidOperationException("Use PdfPageNumber inside PdfHeader or PdfFooter content.");
        PdfStandardFont? standard = null;
        string? family = Font;
        if (Font != null && Enum.TryParse(Font, true, out PdfStandardFont parsed)) { standard = parsed; family = null; }
        var style = new PdfTextRun(string.Empty, bold: Bold, italic: Italic, strike: Strikethrough,
            color: Color, backgroundColor: BackgroundColor, fontSize: Size, fontFamily: family, standardFont: standard,
            baseline: Baseline ?? PdfTextBaseline.Normal, underlineStyle: Underline ?? OfficeTextDecorationStyle.None);
        running.CurrentPage(style);
        if (IncludeTotal) { running.Run(new PdfTextRun(Separator ?? "/")); running.TotalPages(style); }
    }
}
