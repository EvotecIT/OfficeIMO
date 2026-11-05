using OfficeIMO;
using PSWriteOffice.Foundation;
using System.Management.Automation;
using OfficeIMO.Visio.Stencils;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Exports preview artwork from a Visio stencil package into a browsable HTML gallery.</summary>
/// <example>
///   <summary>Create a stencil preview gallery.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$gallery = Export-OfficeVisioStencilPreviewGallery -Path .\MyShapes.vssx -OutputDirectory .\StencilGallery -Title 'Custom stencil previews' -PassThru
/// $gallery | Select-Object PackagePath, IndexPath, BrowserRenderableCount, ThumbnailCount</code>
///   <para>Extracts preview artwork from package-backed masters and writes preview files plus an HTML index.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeVisioStencilPreviewGallery", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
[OutputType(typeof(VisioStencilPreviewGallery))]
public sealed class ExportOfficeVisioStencilPreviewGalleryCommand : PSCmdlet
{
    /// <summary>Visio package path, such as .vsdx, .vssx, or .vstx.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Literal Visio package path.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath")]
    public string LiteralPath { get; set; } = string.Empty;

    /// <summary>Directory that receives preview payloads, thumbnails, and the HTML index.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputDirectory { get; set; } = string.Empty;

    /// <summary>Gallery title. When omitted, a title is derived from the package name.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Optional master filters for package-backed catalogs.</summary>
    [Parameter]
    public string[]? MasterName { get; set; }

    /// <summary>Include unsupported package masters when looking for preview artwork.</summary>
    [Parameter]
    public SwitchParameter IncludeUnsupportedMasters { get; set; }

    /// <summary>Skip reading master dimensions from package master parts.</summary>
    [Parameter]
    public SwitchParameter NoLearnMasterDimensions { get; set; }

    /// <summary>Subdirectory that receives extracted preview payload files.</summary>
    [Parameter]
    public string PreviewDirectoryName { get; set; } = "previews";

    /// <summary>Generated HTML index file name.</summary>
    [Parameter]
    public string IndexFileName { get; set; } = "index.html";

    /// <summary>Do not write an HTML index file.</summary>
    [Parameter]
    public SwitchParameter NoIndex { get; set; }

    /// <summary>Do not write browser-renderable thumbnail wrappers.</summary>
    [Parameter]
    public SwitchParameter NoThumbnails { get; set; }

    /// <summary>Subdirectory that receives generated thumbnail wrappers.</summary>
    [Parameter]
    public string ThumbnailDirectoryName { get; set; } = "thumbnails";

    /// <summary>Generated thumbnail width in pixels.</summary>
    [Parameter]
    public int ThumbnailWidth { get; set; } = 220;

    /// <summary>Generated thumbnail height in pixels.</summary>
    [Parameter]
    public int ThumbnailHeight { get; set; } = 160;

    /// <summary>Default width for package-backed stencils when dimensions cannot be learned.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength DefaultWidth { get; set; } = OfficeLength.Inches(1.8);

    /// <summary>Default height for package-backed stencils when dimensions cannot be learned.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength DefaultHeight { get; set; } = OfficeLength.Inches(0.9);

    /// <summary>Emit the structured gallery result after the files are written.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sources = OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath");
        if (sources.Count != 1) throw new PSArgumentException("A stencil gallery requires one package.");
        var path = sources[0];
        var outputDirectory = VisioCommandUtilities.ResolvePath(this, OutputDirectory);
        if (!ShouldProcess(outputDirectory, "Export Visio stencil preview gallery"))
        {
            return;
        }

        var packageOptions = VisioStencilCommandUtilities.BuildPackageLoadOptions(
            null,
            null,
            null,
            MasterName,
            IncludeUnsupportedMasters.IsPresent,
            NoLearnMasterDimensions.IsPresent,
            noPreviewImageMetadata: false,
            noConnectionPointMetadata: true,
            DefaultWidth.ToInches(),
            DefaultHeight.ToInches());

        var galleryOptions = new VisioStencilPreviewGalleryOptions
        {
            Title = Title,
            PreviewDirectoryName = PreviewDirectoryName,
            IndexFileName = IndexFileName,
            WriteIndex = !NoIndex.IsPresent,
            WriteBrowserRenderableThumbnails = !NoThumbnails.IsPresent,
            ThumbnailDirectoryName = ThumbnailDirectoryName,
            ThumbnailWidth = ThumbnailWidth,
            ThumbnailHeight = ThumbnailHeight
        };

        var result = VisioStencilPackageCatalog.CreatePreviewGallery(path, outputDirectory, packageOptions, galleryOptions);
        if (PassThru.IsPresent)
        {
            WriteObject(result);
        }
    }
}
