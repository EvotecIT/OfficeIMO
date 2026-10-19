using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates literal text decoding and PDF layout settings for Export-OfficeDocumentPdf.</summary>
/// <example><summary>Read UTF-16 text with four-column tabs.</summary><prefix>PS&gt; </prefix>
/// <code>$options = New-OfficeTextPdfOptions -EncodingName utf-16 -TabSize 4
/// Export-OfficeDocumentPdf -InputPath .\Report.txt -Path .\Report.pdf -TextOptions $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeTextPdfOptions")]
[OutputType(typeof(PdfPlainTextOptions))]
public sealed class NewOfficeTextPdfOptionsCommand : PSCmdlet {
    /// <summary>Explicit encoding; otherwise use a Unicode BOM or strict UTF-8.</summary>
    [Parameter] public string? EncodingName { get; set; }
    /// <summary>Source columns per tab stop.</summary>
    [Parameter] [ValidateRange(1, 32)] public int TabSize { get; set; } = 8;
    /// <summary>Maximum decoded and expanded characters.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaximumCharacters { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum generated pages.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int MaximumPages { get; set; } = 10_000;
    /// <summary>PDF fonts and page geometry; default is ten-point Courier.</summary>
    [Parameter] public PdfOptions? PdfOptions { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfPlainTextOptions { EncodingName = EncodingName, TabSize = TabSize,
            MaximumCharacters = MaximumCharacters, MaximumPages = MaximumPages };
        if (PdfOptions != null) options.PdfOptions = PdfOptions.Clone();
        WriteObject(options.Clone());
    }
}
