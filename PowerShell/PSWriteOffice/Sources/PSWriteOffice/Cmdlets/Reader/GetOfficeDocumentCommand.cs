using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads a supported file into the OfficeIMO shared document read result envelope.</summary>
/// <remarks>
/// Use <c>-AsJson</c> for a deterministic serialized payload suitable for indexing or API handoff.
/// </remarks>
/// <example>
///   <summary>Read a document envelope as JSON.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$json = Get-OfficeDocument -Path .\Report.pdf -AsJson -Indented
/// $json | Set-Content -Path .\Report.reader.json</code>
///   <para>Reads a supported document and stores the normalized Reader result for indexing or diagnostics.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocument")]
[Alias("Read-OfficeDocument")]
[OutputType(typeof(OfficeDocumentReadResult), typeof(string))]
public sealed class GetOfficeDocumentCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return the read result as JSON.</summary>
    [Parameter]
    public SwitchParameter AsJson { get; set; }

    /// <summary>Indent JSON output.</summary>
    [Parameter]
    public SwitchParameter Indented { get; set; }

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

    /// <summary>
    /// Compute Word page locations and reconstruct RTF pages from explicit page and section breaks.
    /// PDF and logical-container formats expose their native page-like locations without this switch.
    /// </summary>
    [Parameter]
    public SwitchParameter IncludePageLocations { get; set; }

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from the store. The default is 1,000.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from an email store.</summary>
    [Parameter]
    public SwitchParameter AllStoreItems { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = ReaderCommandUtilities.ResolvePath(this, Path);
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
            IncludePageLocations.IsPresent,
            ResolveStoreItemLimit(MaxStoreItems, AllStoreItems));

        var reader = ResolveReader(configuration.HandlerOptions);
        WriteObject(AsJson.IsPresent
            ? reader.ReadDocumentJson(path, configuration.ReaderOptions, Indented.IsPresent)
            : reader.ReadDocument(path, configuration.ReaderOptions));
    }
}
