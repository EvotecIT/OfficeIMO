using System.IO;
using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads structured tables discovered by OfficeIMO.Reader from a supported document.</summary>
/// <example>
///   <summary>Export discovered tables as sidecars.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$tables = Get-OfficeDocumentTable -Path .\Workbook.xlsx -AsExport -OutputDirectory .\reader-tables -Indented
/// $tables | Select-Object Id, CsvPath, JsonPath</code>
///   <para>Reads tables from a supported document and writes deterministic CSV, Markdown, and JSON sidecars.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentTable")]
[Alias("Read-OfficeDocumentTable")]
[OutputType(typeof(ReaderTable), typeof(ReaderTableExportBundle), typeof(ReaderTableMaterializedExport))]
public sealed class GetOfficeDocumentTableCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return deterministic CSV, Markdown, and JSON export bundles instead of table models.</summary>
    [Parameter]
    public SwitchParameter AsExport { get; set; }

    /// <summary>Optional directory where table sidecars should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Do not overwrite existing sidecar files.</summary>
    [Parameter]
    public SwitchParameter NoOverwrite { get; set; }

    /// <summary>Indent JSON payloads in export bundles and sidecars.</summary>
    [Parameter]
    public SwitchParameter Indented { get; set; }

    /// <summary>Do not write CSV sidecars when <c>-OutputDirectory</c> is used.</summary>
    [Parameter]
    public SwitchParameter NoCsv { get; set; }

    /// <summary>Do not write Markdown sidecars when <c>-OutputDirectory</c> is used.</summary>
    [Parameter]
    public SwitchParameter NoMarkdown { get; set; }

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
            var exports = reader.ReadTableExports(path, configuration.ReaderOptions, Indented.IsPresent);
            var materialized = exports.WriteTableExportsToDirectory(outputDirectory, new ReaderTableExportMaterializationOptions
            {
                Overwrite = !NoOverwrite.IsPresent,
                IncludeCsv = !NoCsv.IsPresent,
                IncludeMarkdown = !NoMarkdown.IsPresent,
                IncludeJson = !NoJson.IsPresent
            });
            WriteObject(materialized, enumerateCollection: true);
            return;
        }

        if (AsExport.IsPresent)
        {
            WriteObject(reader.ReadTableExports(path, configuration.ReaderOptions, Indented.IsPresent), enumerateCollection: true);
            return;
        }

        WriteObject(reader.ReadTables(path, configuration.ReaderOptions), enumerateCollection: true);
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
