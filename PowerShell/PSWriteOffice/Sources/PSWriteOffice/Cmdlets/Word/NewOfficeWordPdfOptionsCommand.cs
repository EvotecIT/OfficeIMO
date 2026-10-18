using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates discoverable Word-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Configure metadata, page numbers, and font embedding.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeWordPdfOptions -Title 'Service report' -Author 'Evotec' -IncludePageNumbers -AllowSystemFontEmbedding
/// Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -WordOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWordPdfOptions")]
[OutputType(typeof(WordToPdfOptions))]
public sealed class NewOfficeWordPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Default font family used when the document does not specify one.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>PDF page size.</summary>
    [Parameter]
    public PageSize? PageSize { get; set; }

    /// <summary>PDF page orientation.</summary>
    [Parameter]
    public OfficePageOrientation? Orientation { get; set; }

    /// <summary>Fallback Word page size for sections without page settings.</summary>
    [Parameter]
    public WordPageSize? DefaultPageSize { get; set; }

    /// <summary>Fallback page orientation for sections without page settings.</summary>
    [Parameter]
    public OfficePageOrientation? DefaultOrientation { get; set; }

    /// <summary>Left page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginLeft { get; set; }

    /// <summary>Top page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginTop { get; set; }

    /// <summary>Right page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginRight { get; set; }

    /// <summary>Bottom page margin in PDF points.</summary>
    [Parameter]
    [ValidateRange(0d, double.MaxValue)]
    public double? MarginBottom { get; set; }

    /// <summary>PDF title metadata.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>PDF author metadata.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>PDF subject metadata.</summary>
    [Parameter]
    public string? Subject { get; set; }

    /// <summary>PDF keywords metadata.</summary>
    [Parameter]
    public string? Keywords { get; set; }

    /// <summary>Include page numbers in the generated PDF.</summary>
    [Parameter]
    public SwitchParameter IncludePageNumbers { get; set; }

    /// <summary>Page number text format.</summary>
    [Parameter]
    public string? PageNumberFormat { get; set; }

    /// <summary>Draw default borders for tables that do not specify borders.</summary>
    [Parameter]
    public SwitchParameter DefaultTableBorders { get; set; }

    /// <summary>Allow embedding fonts discovered on the current system.</summary>
    [Parameter]
    public SwitchParameter AllowSystemFontEmbedding { get; set; }

    /// <summary>Allow embedding installed fonts selected by a trusted Word document. System-font embedding must also be enabled.</summary>
    [Parameter]
    public SwitchParameter AllowDocumentFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new WordToPdfOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        if (!string.IsNullOrWhiteSpace(FontFamily)) options.FontFamily = FontFamily;
        if (PageSize.HasValue) options.PageSize = PageSize.Value;
        if (Orientation.HasValue) options.Orientation = Orientation.Value;
        if (DefaultPageSize.HasValue) options.DefaultPageSize = DefaultPageSize.Value;
        if (DefaultOrientation.HasValue) options.DefaultOrientation = DefaultOrientation.Value;
        if (HasMargins()) {
            PageMargins defaults = PageMargins.Normal;
            options.Margins = new PageMargins(
                MarginLeft ?? defaults.Left,
                MarginTop ?? defaults.Top,
                MarginRight ?? defaults.Right,
                MarginBottom ?? defaults.Bottom);
        }
        if (!string.IsNullOrWhiteSpace(Title)) options.Title = Title;
        if (!string.IsNullOrWhiteSpace(Author)) options.Author = Author;
        if (!string.IsNullOrWhiteSpace(Subject)) options.Subject = Subject;
        if (!string.IsNullOrWhiteSpace(Keywords)) options.Keywords = Keywords;
        if (IsBound(nameof(IncludePageNumbers))) options.IncludePageNumbers = IncludePageNumbers.IsPresent;
        if (!string.IsNullOrWhiteSpace(PageNumberFormat)) options.PageNumberFormat = PageNumberFormat;
        if (IsBound(nameof(DefaultTableBorders))) options.DefaultTableBorders = DefaultTableBorders.IsPresent;
        if (IsBound(nameof(AllowSystemFontEmbedding))) options.ResourcePolicy.AllowSystemFontEmbedding = AllowSystemFontEmbedding.IsPresent;
        if (IsBound(nameof(AllowDocumentFontEmbedding))) options.ResourcePolicy.AllowDocumentFontEmbedding = AllowDocumentFontEmbedding.IsPresent;
        WriteObject(options);
    }

    private bool HasMargins() => MarginLeft.HasValue || MarginTop.HasValue || MarginRight.HasValue || MarginBottom.HasValue;
    private bool IsBound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
}
