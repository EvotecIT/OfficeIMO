using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a row of side-by-side PdfColumn blocks.</summary>
/// <example><summary>Use the native OfficePdfRow command.</summary><code>New-OfficePdf -Path .\TwoColumns.pdf {
///     PdfRow {
///         PdfColumn { PdfParagraph 'Left column' }
///         PdfColumn { PdfParagraph 'Right column' }
///     }
/// }</code></example>
[Cmdlet("Add", "OfficePdfRow", DefaultParameterSetName = "Context")]
[OutputType(typeof(PdfDocument))]
[Alias("PdfRow") ]
public sealed class AddOfficePdfRowCommand : OfficeCmdlet {
    /// <summary>Content for this operation.</summary>
    [Parameter(ParameterSetName = "Context", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Document", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }
    /// <summary>Gap for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? Gap { get; set; }
    /// <summary>SpacingBefore for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? SpacingBefore { get; set; }
    /// <summary>SpacingAfter for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? SpacingAfter { get; set; }
    /// <summary>KeepTogether for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter KeepTogether { get; set; }
    /// <summary>KeepWithNext for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter KeepWithNext { get; set; }
    /// <summary>ColumnSeparatorColor for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? ColumnSeparatorColor { get; set; }
    /// <summary>ColumnSeparatorWidth for this operation.</summary>
    [Parameter(ParameterSetName = "Context")]
    [Parameter(ParameterSetName = "Document")]
    [OfficeLengthTransformation]
    public OfficeLength? ColumnSeparatorWidth { get; set; }
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", Mandatory = true, ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }
    /// <summary>Document set only; inside a block the document is returned by New-OfficePdf.</summary>
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var columns = new List<PdfColumnContent>();
        using (OfficeDocumentContext.Push(columns)) Content!.InvokeReturnAsIs();
        var document = PdfCommandUtilities.ComposeContent(this, Document, Document == null ? "Context" : "Document", "Document", content => content.Row(row => {
            if (Gap.HasValue) row.Gap(Gap.Value);
            row.Style(new PdfRowStyle {
                SpacingBefore = SpacingBefore ?? OfficeLength.Zero,
                SpacingAfter = SpacingAfter ?? OfficeLength.Zero,
                KeepTogether = KeepTogether,
                KeepWithNext = KeepWithNext,
                ColumnSeparatorColor = ColumnSeparatorColor,
                ColumnSeparatorWidth = ColumnSeparatorWidth ?? OfficeLength.Points(0.5)
            });
            foreach (var column in columns) row.Column(column.Width, column.Apply);
        }));
        if (PassThru && document != null) WriteObject(document);
    }
}
