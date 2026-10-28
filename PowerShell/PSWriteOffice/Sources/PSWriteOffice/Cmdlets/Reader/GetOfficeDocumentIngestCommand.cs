using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads a folder into an OfficeIMO.Reader ingestion summary.</summary>
/// <example>
///   <summary>Ingest a report folder.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ingest = Get-OfficeDocumentIngest -FolderPath .\Reports -Extension docx,pdf,rtf -MaxFiles 50
/// $ingest.Files | Select-Object Path, Parsed, ChunksProduced</code>
///   <para>Reads supported files from a folder and returns the ingestion summary with per-file status and chunk counts.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentIngest")]
[OutputType(typeof(ReaderIngestResult))]
public sealed class GetOfficeDocumentIngestCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>Folder path to ingest.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string FolderPath { get; set; } = string.Empty;

    /// <summary>Do not recurse into child folders.</summary>
    [Parameter]
    public SwitchParameter NoRecurse { get; set; }

    /// <summary>Maximum number of folder files to read.</summary>
    [Parameter]
    public int? MaxFiles { get; set; }

    /// <summary>Maximum total folder bytes to read.</summary>
    [Parameter]
    public long? MaxTotalBytes { get; set; }

    /// <summary>Allowed folder extensions such as .docx, .xlsx, .pdf, .html, or json.</summary>
    [Parameter]
    [Alias("Extensions")]
    public string[]? Extension { get; set; }

    /// <summary>Do not materialize chunks in the returned ingestion result.</summary>
    [Parameter]
    public SwitchParameter NoChunks { get; set; }

    /// <summary>Maximum input size in bytes.</summary>
    [Parameter]
    public long? MaxInputBytes { get; set; }

    /// <summary>OpenXML maximum characters per part.</summary>
    [Parameter]
    public long? OpenXmlMaxCharactersInPart { get; set; }

    /// <summary>Maximum emitted chunk characters.</summary>
    [Parameter]
    public int? MaxChars { get; set; }

    /// <summary>Maximum table rows per emitted table chunk.</summary>
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

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from each email store.</summary>
    [Parameter]
    public SwitchParameter AllStoreItems { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var folderPath = ReaderCommandUtilities.ResolvePath(this, FolderPath);
        var folderOptions = ReaderCommandUtilities.BuildFolderOptions(!NoRecurse.IsPresent, MaxFiles, MaxTotalBytes, Extension);
        var configuration = ReaderCommandUtilities.BuildReadConfiguration(
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
            !NoHashes.IsPresent,
            includePageLocations: false,
            maxStoreItems: ResolveStoreItemLimit(MaxStoreItems, AllStoreItems));

        var reader = ResolveReader(configuration.HandlerOptions);
        WriteObject(reader.ReadFolderDetailed(
            folderPath,
            folderOptions,
            configuration.ReaderOptions,
            includeChunks: !NoChunks.IsPresent));
    }
}
