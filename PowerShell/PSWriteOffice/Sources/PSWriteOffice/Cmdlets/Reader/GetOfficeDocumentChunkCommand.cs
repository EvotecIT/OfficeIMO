using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads supported Office, PDF, Markdown, RTF, HTML, CSV, JSON, XML, YAML, ZIP, EPUB, Visio, and text files into normalized OfficeIMO.Reader chunks.</summary>
/// <remarks>
/// This is a thin adapter over <see cref="OfficeDocumentReader"/>. The OfficeIMO.Reader engine owns detection,
/// extraction, hashing, and chunk shaping.
/// </remarks>
/// <example>
///   <summary>Read semantic chunks from an RTF document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeRtf -Path .\Report.rtf -Text 'Summary', 'Ready for review'
/// Get-OfficeDocumentChunk -Path .\Report.rtf | Select-Object Kind, Text</code>
///   <para>Creates a small RTF file and reads it back through the Reader adapter as normalized chunks.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentChunk", DefaultParameterSetName = FileParameterSet)]
[Alias("Read-OfficeDocumentChunk")]
[OutputType(typeof(ReaderChunk))]
public sealed class GetOfficeDocumentChunkCommand : OfficeDocumentReaderCommandBase
{
    private const string FileParameterSet = "File";
    private const string FolderParameterSet = "Folder";

    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = FileParameterSet)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Folder path to read.</summary>
    [Parameter(Mandatory = true, ParameterSetName = FolderParameterSet)]
    public string FolderPath { get; set; } = string.Empty;

    /// <summary>Do not recurse into child folders when reading a folder.</summary>
    [Parameter(ParameterSetName = FolderParameterSet)]
    public SwitchParameter NoRecurse { get; set; }

    /// <summary>Maximum number of folder files to read.</summary>
    [Parameter(ParameterSetName = FolderParameterSet)]
    public int? MaxFiles { get; set; }

    /// <summary>Maximum total folder bytes to read.</summary>
    [Parameter(ParameterSetName = FolderParameterSet)]
    public long? MaxTotalBytes { get; set; }

    /// <summary>Allowed folder extensions such as .docx, .xlsx, .pdf, or md.</summary>
    [Parameter(ParameterSetName = FolderParameterSet)]
    [Alias("Extensions")]
    public string[]? Extension { get; set; }

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
        var configuration = BuildOptions();
        var reader = ResolveReader(configuration.HandlerOptions);

        if (ParameterSetName == FolderParameterSet)
        {
            var folderPath = ReaderCommandUtilities.ResolvePath(this, FolderPath);
            var folderOptions = ReaderCommandUtilities.BuildFolderOptions(!NoRecurse.IsPresent, MaxFiles, MaxTotalBytes, Extension);
            foreach (var chunk in reader.ReadFolder(folderPath, folderOptions, configuration.ReaderOptions))
            {
                WriteObject(chunk);
            }

            return;
        }

        var path = ReaderCommandUtilities.ResolvePath(this, Path);
        foreach (var chunk in reader.Read(path, configuration.ReaderOptions))
        {
            WriteObject(chunk);
        }
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
            !NoHashes.IsPresent,
            includePageLocations: false,
            maxStoreItems: ResolveStoreItemLimit(MaxStoreItems, AllStoreItems));
    }
}
