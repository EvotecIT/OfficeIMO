using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Exports one or more Word pages through the format-neutral OfficeIMO image pipeline.</summary>
/// <example>
///   <summary>Export the first page as SVG.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Report.svg -Format Svg</code>
///   <para>Writes the image quietly. Add <c>-PassThru</c> to receive the structured export result.</para>
/// </example>
/// <example>
///   <summary>Export every page as JPEG files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeWordImage -Path .\Report.docx -OutputPath .\Pages -Format Jpeg -AllPages</code>
///   <para>For a bounded batch, create options with <c>New-OfficeWordImageOptions -PageIndex 0 -PageCount 2</c>.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeWordImage", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficeWordImageCommand : PSCmdlet
{
    /// <summary>Path to the Word document.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Open Word document instance.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document")]
    public WordDocument Document { get; set; } = null!;

    /// <summary>Destination image file, or destination folder when -AllPages or Options.PageCount requests a batch.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional page, size, scale, theme, and rendering settings.</summary>
    [Parameter]
    public WordImageExportOptions? Options { get; set; }

    /// <summary>Export every estimated page to the destination folder.</summary>
    [Parameter]
    public SwitchParameter AllPages { get; set; }

    /// <summary>Emit the structured image export result.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        bool batch = AllPages.IsPresent || Options?.PageCount.HasValue == true;
        if (!ShouldProcess(output, batch ? $"Export Word pages as {Format}" : $"Export Word page as {Format}")) return;
        if (batch)
        {
            Directory.CreateDirectory(output);
        }
        else
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        }
        WordDocument? owned = null;
        try
        {
            var document = Document;
            if (ParameterSetName == "Path")
            {
                var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                owned = WordDocumentService.LoadDocument(input, readOnly: true, autoSave: false);
                document = owned;
            }

            WordImageExportOptions effectiveOptions = Options?.Clone() ?? new WordImageExportOptions();
            if (AllPages.IsPresent)
            {
                effectiveOptions.PageIndex = 0;
                effectiveOptions.PageCount = null;
            }

            if (batch)
            {
                IReadOnlyList<OfficeImageExportResult> results = document.SaveAsImages(output, Format, effectiveOptions);
                if (PassThru.IsPresent) WriteObject(results, enumerateCollection: true);
            }
            else
            {
                OfficeImageExportResult result = document.ExportImage(Format, effectiveOptions).Save(output);
                if (PassThru.IsPresent) WriteObject(result);
            }
        }
        finally
        {
            owned?.Dispose();
        }
    }
}
