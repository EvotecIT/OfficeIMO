using System.Globalization;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports PDF pages through the shared PNG, JPEG, TIFF, SVG, or WebP image contract.</summary>
/// <example>
///   <summary>Export selected pages as PNG files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\Pages -PageRange '1-3,5'</code>
///   <para>Writes the selected pages. Add <c>-PassThru</c> to receive normalized image results with rendering diagnostics.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficePdfImage", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficePdfImageCommand : PSCmdlet
{
    /// <summary>Path to the PDF.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination folder.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional one-based ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional DPI, scale, thumbnail, encoding, diagnostics, and resource limits.</summary>
    [Parameter]
    public PdfImageExportOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing settings.</summary>
    [Parameter]
    public PdfLoadOptions? ReadOptions { get; set; }

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Emit one structured image export result per saved page.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var options = Options ?? new PdfImageExportOptions();
        var readOptions = PdfCommandUtilities.CreateReadOptions(
            ReadOptions,
            Password,
            IgnorePermissionRestrictions.IsPresent);
        var selection = string.IsNullOrWhiteSpace(PageRange) ? null : PdfPageSelection.Parse(PageRange!);
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export PDF pages as {Format}")) return;
        Directory.CreateDirectory(output);
        var document = PdfCommandUtilities.LoadDocument(input, readOptions);
        var pages = document.Render.ExportImages(Format, options, selection);
        for (int index = 0; index < pages.Count; index++)
        {
            var page = pages[index];
            int pageNumber = GetPageNumber(page, index + 1);
            var file = System.IO.Path.Combine(output, $"page-{pageNumber:D4}{page.FileExtension}");
            OfficeImageExportResult result = page.Save(file, OfficeImageExportFileConflictPolicy.Replace);
            if (PassThru.IsPresent) WriteObject(result);
        }
    }

    private static int GetPageNumber(OfficeImageExportResult result, int fallback)
    {
        const string prefix = "Page ";
        return result.Name != null &&
               result.Name.StartsWith(prefix, System.StringComparison.Ordinal) &&
               int.TryParse(result.Name.Substring(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out int pageNumber)
            ? pageNumber
            : fallback;
    }
}
