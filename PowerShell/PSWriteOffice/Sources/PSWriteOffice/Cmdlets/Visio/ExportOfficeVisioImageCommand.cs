using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using PSWriteOffice.Services.Visio;

namespace PSWriteOffice.Cmdlets.Visio;

/// <summary>Exports selected Visio pages through the format-neutral OfficeIMO image pipeline.</summary>
/// <example>
///   <summary>Export every page as PNG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeVisioImage -Path .\diagram.vsdx -OutputPath .\Images -Format Png</code>
///   <para>Writes one PNG per selected page. Add <c>-PassThru</c> to receive one result object per file.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeVisioImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeVisioImageCommand : PSCmdlet
{
    /// <summary>Path to a Visio document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Visio document instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public VisioDocument Document { get; set; } = null!;

    /// <summary>Destination folder.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional page selection, size, concurrency, and rendering settings.</summary>
    [Parameter]
    public VisioImageExportOptions? Options { get; set; }

    /// <summary>Emit one structured image export result per saved page.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = VisioCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(output, $"Export Visio pages as {Format}"))
        {
            return;
        }

        Directory.CreateDirectory(output);
        var document = VisioCommandUtilities.ResolveDocument(this, Document, ParameterSetName == "Path" ? Path : null);
        IReadOnlyList<OfficeImageExportResult> results = document.SaveAsImages(output, Format, Options);
        if (PassThru.IsPresent) WriteObject(results, enumerateCollection: true);
    }
}
