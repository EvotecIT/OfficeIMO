using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Exports a Visio document page to native dependency-free PNG.</summary>
/// <example>
///   <summary>Export a diagram preview to PNG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\ServiceMap.vsdx { VisioRectangle -Text 'API' -X 2 -Y 4 }
/// ConvertTo-OfficeVisioPng -Path .\ServiceMap.vsdx -OutputPath .\ServiceMap.png -PixelsPerInch 144</code>
///   <para>Creates a diagram and exports the first page to a PNG preview.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeVisioPng", DefaultParameterSetName = PathParameterSet, SupportsShouldProcess = true)]
[Alias("ConvertTo-VisioPng")]
[OutputType(typeof(byte[]), typeof(FileInfo))]
public sealed class ConvertToOfficeVisioPngCommand : PSCmdlet {
    private const string PathParameterSet = "Path";
    private const string DocumentParameterSet = "Document";

    /// <summary>Visio .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = PathParameterSet)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Visio document object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = DocumentParameterSet)]
    public VisioDocument? Document { get; set; }

    /// <summary>Optional output PNG path.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Zero-based page index to export.</summary>
    [Parameter]
    public int PageIndex { get; set; }

    /// <summary>PNG pixels per Visio inch.</summary>
    [Parameter]
    public double? PixelsPerInch { get; set; }

    /// <summary>Background color name or hex value. Use -Transparent for transparent output.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Use transparent PNG background.</summary>
    [Parameter]
    public SwitchParameter Transparent { get; set; }

    /// <summary>Do not render shape text.</summary>
    [Parameter]
    public SwitchParameter NoText { get; set; }

    /// <summary>Optional TrueType/OpenType font file used for text outlines.</summary>
    [Parameter]
    public string? FontFilePath { get; set; }

    /// <summary>Optional font face name used when selecting from a font collection.</summary>
    [Parameter]
    public string? FontFaceName { get; set; }

    /// <summary>Optional zero-based font collection index.</summary>
    [Parameter]
    public int? FontCollectionIndex { get; set; }

    /// <summary>Do not render built-in stencil artwork.</summary>
    [Parameter]
    public SwitchParameter NoStencilArtwork { get; set; }

    /// <summary>Do not render connector labels.</summary>
    [Parameter]
    public SwitchParameter NoConnectorLabels { get; set; }

    /// <summary>Do not resolve connector label overlaps at export time.</summary>
    [Parameter]
    public SwitchParameter NoConnectorLabelOverlapResolution { get; set; }

    /// <summary>Supersampling factor for smoother raster output.</summary>
    [Parameter]
    public int? Supersampling { get; set; }

    /// <summary>Open the PNG after saving.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = VisioCommandUtilities.BuildImageOptions(
            this,
            PageIndex,
            PixelsPerInch,
            BackgroundColor,
            Transparent.IsPresent,
            NoText.IsPresent,
            FontFilePath,
            FontFaceName,
            FontCollectionIndex,
            NoStencilArtwork.IsPresent,
            NoConnectorLabels.IsPresent,
            NoConnectorLabelOverlapResolution.IsPresent,
            Supersampling,
            includeSvgXmlDeclaration: false);

        string? targetPath = null;

        if (!string.IsNullOrWhiteSpace(OutputPath)) {
            targetPath = VisioCommandUtilities.ResolveImageOutputPath(
                this,
                OutputPath!,
                OfficeImageExportFormat.Png);
            if (!ShouldProcess(targetPath, "Write Visio PNG")) {
                return;
            }
        }

        var document = VisioCommandUtilities.ResolveDocument(this, Document, Path);
        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Png, options);

        if (targetPath != null) {
            VisioCommandUtilities.EnsureDirectory(targetPath);
            OfficeImageExportResult saved = result.Save(targetPath);

            if (Open.IsPresent) {
                FileOpenService.Open(saved.SavedPath!);
            }

            WriteObject(new FileInfo(saved.SavedPath!));
            return;
        }

        WriteObject(result.Bytes, enumerateCollection: false);
    }
}
