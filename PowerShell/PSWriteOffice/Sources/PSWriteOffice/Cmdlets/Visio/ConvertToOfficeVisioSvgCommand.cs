using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Exports a Visio document page to dependency-free SVG.</summary>
/// <example>
///   <summary>Export a diagram to SVG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>OfficeVisio -Path .\ServiceMap.vsdx { VisioRectangle -Text 'API' -X 2 -Y 4 }
/// ConvertTo-OfficeVisioSvg -Path .\ServiceMap.vsdx -OutputPath .\ServiceMap.svg -Transparent</code>
///   <para>Creates a diagram and exports the first page to dependency-free SVG.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeVisioSvg", DefaultParameterSetName = PathParameterSet, SupportsShouldProcess = true)]
[Alias("ConvertTo-VisioSvg")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeVisioSvgCommand : PSCmdlet {
    private const string PathParameterSet = "Path";
    private const string DocumentParameterSet = "Document";

    /// <summary>Visio .vsdx path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = PathParameterSet)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Visio document object.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = DocumentParameterSet)]
    public VisioDocument? Document { get; set; }

    /// <summary>Optional output SVG path.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Zero-based page index to export.</summary>
    [Parameter]
    public int PageIndex { get; set; }

    /// <summary>SVG pixels per Visio inch.</summary>
    [Parameter]
    public double? PixelsPerInch { get; set; }

    /// <summary>Background color name or hex value. Use -Transparent for transparent output.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? BackgroundColor { get; set; }

    /// <summary>Use transparent SVG background.</summary>
    [Parameter]
    public SwitchParameter Transparent { get; set; }

    /// <summary>Do not render shape text.</summary>
    [Parameter]
    public SwitchParameter NoText { get; set; }

    /// <summary>Do not render built-in stencil artwork.</summary>
    [Parameter]
    public SwitchParameter NoStencilArtwork { get; set; }

    /// <summary>Do not render connector labels.</summary>
    [Parameter]
    public SwitchParameter NoConnectorLabels { get; set; }

    /// <summary>Do not resolve connector label overlaps at export time.</summary>
    [Parameter]
    public SwitchParameter NoConnectorLabelOverlapResolution { get; set; }

    /// <summary>Include XML declaration in the generated SVG.</summary>
    [Parameter]
    public SwitchParameter IncludeXmlDeclaration { get; set; }

    /// <summary>Open the SVG after saving.</summary>
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
            fontFilePath: null,
            fontFaceName: null,
            fontCollectionIndex: null,
            NoStencilArtwork.IsPresent,
            NoConnectorLabels.IsPresent,
            NoConnectorLabelOverlapResolution.IsPresent,
            supersampling: null,
            IncludeXmlDeclaration.IsPresent);

        string? targetPath = null;

        if (!string.IsNullOrWhiteSpace(OutputPath)) {
            targetPath = VisioCommandUtilities.ResolveImageOutputPath(
                this,
                OutputPath!,
                OfficeImageExportFormat.Svg);
            if (!ShouldProcess(targetPath, "Write Visio SVG")) {
                return;
            }
        }

        var document = VisioCommandUtilities.ResolveDocument(this, Document, Path);
        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Svg, options);

        if (targetPath != null) {
            VisioCommandUtilities.EnsureDirectory(targetPath);
            OfficeImageExportResult saved = result.Save(targetPath);

            if (Open.IsPresent) {
                FileOpenService.Open(saved.SavedPath!);
            }

            WriteObject(new FileInfo(saved.SavedPath!));
            return;
        }

        WriteObject(Encoding.UTF8.GetString(result.Bytes));
    }
}
