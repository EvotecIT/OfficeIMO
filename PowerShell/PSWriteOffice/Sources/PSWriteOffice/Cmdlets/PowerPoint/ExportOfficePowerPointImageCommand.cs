using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Exports presentation slides as PNG or SVG images with one result per slide.</summary>
/// <example>
///   <summary>Export visible slides as SVG files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficePowerPointImage -Path .\Deck.pptx -OutputPath .\Slides -Format Svg</code>
///   <para>Writes one image per selected slide. Add <c>-PassThru</c> to receive the structured export results.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficePowerPointImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficePowerPointImageCommand : PSCmdlet
{
    /// <summary>Path to the presentation.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open presentation instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Presentation")]
    public PowerPointPresentation Presentation { get; set; } = null!;

    /// <summary>Destination folder.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional slide selection, size, scale, theme, and rendering settings.</summary>
    [Parameter]
    public PowerPointPresentationImageExportOptions? Options { get; set; }

    /// <summary>Emit one structured image export result per saved slide.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export PowerPoint slides as {Format}")) return;
        Directory.CreateDirectory(output);
        PowerPointPresentation? owned = null;
        try
        {
            var presentation = Presentation;
            if (ParameterSetName == "Path")
            {
                var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = PowerPointDocumentService.LoadPresentation(input, readOnly: true);
                presentation = owned;
            }
            IReadOnlyList<OfficeImageExportResult> results = presentation.SaveAsImages(output, Format, Options);
            if (PassThru.IsPresent) WriteObject(results, enumerateCollection: true);
        }
        finally
        {
            if (owned != null)
            {
                PowerPointDocumentService.ClosePresentation(owned, save: false, show: false);
            }
        }
    }
}
