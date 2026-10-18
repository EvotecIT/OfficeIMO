using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace PSWriteOffice.Cmdlets.Html;

/// <summary>Exports an HTML render surface as PNG or SVG with structured diagnostics.</summary>
/// <example>
///   <summary>Render an HTML file to PNG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.png</code>
///   <para>Uses the dependency-free OfficeIMO HTML renderer. Add <c>-PassThru</c> to receive the structured export result.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeHtmlImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeHtmlImageCommand : PSCmdlet
{
    private readonly StringBuilder _pipelineHtml = new();
    private bool _hasPipelineHtml;
    private bool _shouldExport;
    private string _resolvedOutput = string.Empty;

    /// <summary>Path to an HTML file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>HTML markup to render.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Html")]
    public string Html { get; set; } = string.Empty;

    /// <summary>Shared HTML conversion document.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public HtmlConversionDocument Document { get; set; } = null!;

    /// <summary>Destination PNG, JPEG, TIFF, SVG, or WebP path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Zero-based rendered page index.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int PageIndex { get; set; }

    /// <summary>Optional HTML parsing and trust settings for path or markup input.</summary>
    [Parameter]
    public HtmlConversionDocumentOptions? DocumentOptions { get; set; }

    /// <summary>Optional size, pagination, resource, and rendering settings.</summary>
    [Parameter]
    public HtmlRenderOptions? RenderOptions { get; set; }

    /// <summary>Emit the structured image export result.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        _resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        _shouldExport = ShouldProcess(_resolvedOutput, $"Export HTML page as {Format}");
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!_shouldExport) return;

        if (ParameterSetName == "Html")
        {
            if (_hasPipelineHtml) _pipelineHtml.Append('\n');
            _pipelineHtml.Append(Html ?? string.Empty);
            _hasPipelineHtml = true;
            return;
        }

        var document = ParameterSetName == "Path"
            ? HtmlConversionDocument.Load(SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), DocumentOptions)
            : Document;
        Export(document);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_shouldExport && ParameterSetName == "Html")
        {
            Export(HtmlConversionDocument.Parse(_pipelineHtml.ToString(), DocumentOptions));
        }
    }

    private void Export(HtmlConversionDocument document)
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(_resolvedOutput) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        OfficeImageExportResult result = document.ExportImage(Format, RenderOptions, PageIndex).Save(_resolvedOutput);
        if (PassThru.IsPresent) WriteObject(result);
    }
}
