using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads visual payloads discovered by OfficeIMO.Reader from a supported document.</summary>
/// <example>
///   <summary>Export discovered visual payloads.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$visuals = Get-OfficeDocumentVisual -Path .\Deck.pptx -AsExport -OutputDirectory .\reader-visuals -Indented
/// $visuals | Select-Object Id, ContentType, PayloadPath</code>
///   <para>Extracts supported visual payloads and writes payload plus metadata sidecars.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentVisual")]
[Alias("Read-OfficeDocumentVisual")]
[OutputType(typeof(ReaderVisual), typeof(ReaderVisualExportBundle), typeof(ReaderVisualMaterializedExport))]
public sealed class GetOfficeDocumentVisualCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return deterministic payload and JSON export bundles instead of visual models.</summary>
    [Parameter]
    public SwitchParameter AsExport { get; set; }

    /// <summary>Optional directory where visual sidecars should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Do not overwrite existing sidecar files.</summary>
    [Parameter]
    public SwitchParameter NoOverwrite { get; set; }

    /// <summary>Indent JSON payloads in export bundles and sidecars.</summary>
    [Parameter]
    public SwitchParameter Indented { get; set; }

    /// <summary>Do not write raw payload sidecars when <c>-OutputDirectory</c> is used.</summary>
    [Parameter]
    public SwitchParameter NoPayload { get; set; }

    /// <summary>Do not write JSON sidecars when <c>-OutputDirectory</c> is used.</summary>
    [Parameter]
    public SwitchParameter NoJson { get; set; }

    /// <summary>Maximum input size in bytes.</summary>
    [Parameter]
    public long? MaxInputBytes { get; set; }

    /// <summary>OpenXML maximum characters per part.</summary>
    [Parameter]
    public long? OpenXmlMaxCharactersInPart { get; set; }

    /// <summary>Maximum emitted chunk characters.</summary>
    [Parameter]
    public int? MaxChars { get; set; }

    /// <summary>Maximum table rows per emitted table.</summary>
    [Parameter]
    public int? MaxTableRows { get; set; }

    /// <summary>Exclude Word footnotes.</summary>
    [Parameter]
    public SwitchParameter ExcludeWordFootnotes { get; set; }

    /// <summary>Exclude PowerPoint speaker notes.</summary>
    [Parameter]
    public SwitchParameter ExcludePowerPointNotes { get; set; }

    /// <summary>Treat the first Excel row as data instead of headers.</summary>
    [Parameter]
    public SwitchParameter NoExcelHeaders { get; set; }

    /// <summary>Excel rows per emitted worksheet chunk.</summary>
    [Parameter]
    public int? ExcelChunkRows { get; set; }

    /// <summary>Optional Excel sheet name to read.</summary>
    [Parameter]
    public string? ExcelSheetName { get; set; }

    /// <summary>Optional Excel A1 range to read.</summary>
    [Parameter]
    public string? ExcelA1Range { get; set; }

    /// <summary>Do not split Markdown by headings.</summary>
    [Parameter]
    public SwitchParameter NoMarkdownHeadingChunks { get; set; }

    /// <summary>Disable source and chunk hash computation.</summary>
    [Parameter]
    public SwitchParameter NoHashes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = ReaderCommandUtilities.ResolvePath(this, Path);
        var configuration = BuildOptions();
        var reader = ResolveReader(configuration.HandlerOptions);

        if (!string.IsNullOrWhiteSpace(OutputDirectory))
        {
            var outputDirectory = ReaderCommandUtilities.ResolvePath(this, OutputDirectory!);
            var exports = reader.ReadVisualExports(path, configuration.ReaderOptions, Indented.IsPresent);
            var materialized = exports.WriteVisualExportsToDirectory(outputDirectory, new ReaderVisualExportMaterializationOptions
            {
                Overwrite = !NoOverwrite.IsPresent,
                IncludePayload = !NoPayload.IsPresent,
                IncludeJson = !NoJson.IsPresent
            });
            WriteObject(materialized, enumerateCollection: true);
            return;
        }

        if (AsExport.IsPresent)
        {
            WriteObject(reader.ReadVisualExports(path, configuration.ReaderOptions, Indented.IsPresent), enumerateCollection: true);
            return;
        }

        WriteObject(reader.ReadVisuals(path, configuration.ReaderOptions), enumerateCollection: true);
    }

    private ReaderCommandConfiguration BuildOptions()
    {
        return ReaderCommandUtilities.BuildReadConfiguration(
            MaxInputBytes,
            OpenXmlMaxCharactersInPart,
            MaxChars,
            MaxTableRows,
            !ExcludeWordFootnotes.IsPresent,
            !ExcludePowerPointNotes.IsPresent,
            !NoExcelHeaders.IsPresent,
            ExcelChunkRows,
            ExcelSheetName,
            ExcelA1Range,
            !NoMarkdownHeadingChunks.IsPresent,
            !NoHashes.IsPresent);
    }
}
