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

/// <summary>Defines one column of a PdfRow and its content.</summary>
/// <example><summary>Use the native OfficePdfColumn command.</summary><code>New-OfficePdf -Path .\Columns.pdf {
///     PdfRow {
///         PdfColumn { PdfHeading 'Signals' -Level 2 }
///         PdfColumn { PdfPanel 'All checks passed.' }
///     }
/// }</code></example>
[Cmdlet("Add", "OfficePdfColumn", DefaultParameterSetName = "Weight")]
[OutputType(typeof(void))]
[Alias("PdfColumn") ]
public sealed class AddOfficePdfColumnCommand : OfficeCmdlet {
    /// <summary>Content for this operation.</summary>
    [Parameter(ParameterSetName = "Weight", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Width", Mandatory = true, Position = 0)]
    public ScriptBlock? Content { get; set; }
    /// <summary>Weight for this operation.</summary>
    [Parameter(ParameterSetName = "Weight")]
    public double? Weight { get; set; }
    /// <summary>Width for this operation.</summary>
    [Parameter(ParameterSetName = "Width", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var columns = OfficeDocumentContext.Find<List<PdfColumnContent>>() ?? throw new PSInvalidOperationException("Use PdfColumn inside PdfRow.");
        var column = new PdfColumnContent(ParameterSetName == "Width" ? PdfColumnWidth.Fixed(Width) : PdfColumnWidth.Relative(Weight ?? 1D));
        using (OfficeDocumentContext.Push(column.Actions)) Content!.InvokeReturnAsIs();
        columns.Add(column);
    }
}
