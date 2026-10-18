using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads or materializes embedded assets discovered by OfficeIMO.Reader from a supported document.</summary>
/// <example>
///   <summary>List embedded assets in a document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeDocumentAsset -Path .\Report.docx |
///     Select-Object Id, Kind, MediaType, FileName, LengthBytes</code>
///   <para>Reads the normalized asset metadata emitted by the shared Reader envelope.</para>
/// </example>
/// <example>
///   <summary>Export document images to a folder.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeDocumentAsset -Path .\Deck.pptx -OutputDirectory .\reader-assets -Kind image -ValidatePayloadHash</code>
///   <para>Writes materializable image payloads to deterministic filenames and returns materialization results.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentAsset")]
[Alias("Read-OfficeDocumentAsset", "Export-OfficeDocumentAsset")]
[OutputType(typeof(OfficeDocumentAsset), typeof(OfficeDocumentMaterializedAsset))]
public sealed class GetOfficeDocumentAssetCommand : OfficeDocumentReaderCommandBase
{
    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional directory where materializable asset payloads should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Only include assets with one of these normalized kinds, such as image or preview.</summary>
    [Parameter]
    [ArgumentCompleter(typeof(OfficeDocumentAssetKindArgumentCompleter))]
    public string[]? Kind { get; set; }

    /// <summary>Only include assets with one of these media types.</summary>
    [Parameter]
    public string[]? MediaType { get; set; }

    /// <summary>Only include assets with one of these file extensions.</summary>
    [Parameter]
    public string[]? Extension { get; set; }

    /// <summary>Do not overwrite existing payload files when <c>-OutputDirectory</c> is used.</summary>
    [Parameter]
    public SwitchParameter NoOverwrite { get; set; }

    /// <summary>Validate payload hashes before writing assets when hash metadata is present.</summary>
    [Parameter]
    public SwitchParameter ValidatePayloadHash { get; set; }

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
        Func<OfficeDocumentAsset, bool>? predicate = BuildPredicate();

        if (!string.IsNullOrWhiteSpace(OutputDirectory))
        {
            var outputDirectory = ReaderCommandUtilities.ResolvePath(this, OutputDirectory!);
            var result = reader.ReadDocument(path, configuration.ReaderOptions);
            var materialized = result.WriteAssetsToDirectory(outputDirectory, new OfficeDocumentAssetMaterializationOptions
            {
                Overwrite = !NoOverwrite.IsPresent,
                ValidatePayloadHash = ValidatePayloadHash.IsPresent,
                Predicate = predicate
            });
            WriteObject(materialized, enumerateCollection: true);
            return;
        }

        var assets = reader.ReadAssets(path, configuration.ReaderOptions);
        WriteObject(predicate == null ? assets : assets.Where(predicate).ToArray(), enumerateCollection: true);
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

    private Func<OfficeDocumentAsset, bool>? BuildPredicate()
    {
        var kinds = NormalizeFilter(Kind, normalizeExtension: false);
        var mediaTypes = NormalizeFilter(MediaType, normalizeExtension: false);
        var extensions = NormalizeFilter(Extension, normalizeExtension: true);

        if (kinds.Length == 0 && mediaTypes.Length == 0 && extensions.Length == 0)
        {
            return null;
        }

        return asset =>
            Matches(kinds, asset.Kind) &&
            Matches(mediaTypes, asset.MediaType) &&
            Matches(extensions, asset.Extension);
    }

    private static string[] NormalizeFilter(string[]? values, bool normalizeExtension)
    {
        if (values == null || values.Length == 0)
        {
            return Array.Empty<string>();
        }

        return values
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Select(value =>
            {
                var normalized = value.Trim();
                return normalizeExtension && !normalized.StartsWith(".", StringComparison.Ordinal)
                    ? "." + normalized
                    : normalized;
            })
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static bool Matches(string[] filters, string? value)
    {
        return filters.Length == 0 ||
               (!string.IsNullOrWhiteSpace(value) &&
                filters.Any(filter => string.Equals(filter, value, StringComparison.OrdinalIgnoreCase)));
    }
}
