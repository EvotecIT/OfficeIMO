using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using PSWriteOffice.Completion;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Writes document media, tables or visuals using the native Reader materializers.</summary>
/// <example><summary>Extract presentation images and embedded files.</summary>
/// <code>Export-OfficeDocumentAsset ./Review.pptx ./Assets</code></example>
/// <example><summary>Export report tables as CSV files.</summary>
/// <code>Get-ChildItem ./Reports/*.docx | Export-OfficeDocumentAsset -OutputDirectory ./Tables -Type Table -TableFormat Csv</code></example>
/// <example><summary>Keep existing image files.</summary>
/// <code>Export-OfficeDocumentAsset ./Brochure.docx ./Images -Extension png,jpg -NoClobber</code></example>
[Cmdlet(VerbsData.Export, "OfficeDocumentAsset", SupportsShouldProcess = true, DefaultParameterSetName = "Path")]
[OutputType(typeof(OfficeDocumentMaterializedAsset), typeof(ReaderTableMaterializedExport), typeof(ReaderVisualMaterializedExport))]
public sealed class ExportOfficeDocumentAssetCommand : OfficeInputCmdlet {
    private readonly List<string> _paths = new();
    private HashSet<OfficeDocumentAssetType> _types = null!;
    private HashSet<OfficeTableExportFormat> _formats = null!;
    private OfficeReadConfiguration _configuration = null!;
    private OfficeDocumentReader _reader = null!;
    private string _directory = string.Empty;
    private string[] _kinds = Array.Empty<string>(), _mediaTypes = Array.Empty<string>(), _extensions = Array.Empty<string>();

    /// <summary>Destination folder for the native materialized files.</summary>
    [Parameter(Mandatory = true, Position = 1)] [ValidateNotNullOrEmpty]
    public string OutputDirectory { get; set; } = string.Empty;
    /// <summary>Media, tables or visuals. Media is the default.</summary>
    [Parameter] [ValidateNotNullOrEmpty]
    public OfficeDocumentAssetType[] Type { get; set; } = new[] { OfficeDocumentAssetType.Media };
    /// <summary>Selected media asset kinds.</summary>
    [Parameter] [ArgumentCompleter(typeof(OfficeDocumentAssetKindArgumentCompleter))]
    public string[]? Kind { get; set; }
    /// <summary>Selected media MIME types.</summary>
    [Parameter] public string[]? MediaType { get; set; }
    /// <summary>Selected media extensions, with or without a leading dot.</summary>
    [Parameter] public string[]? Extension { get; set; }
    /// <summary>Table representations to write. All formats are included by default.</summary>
    [Parameter] [ValidateNotNullOrEmpty]
    public OfficeTableExportFormat[] TableFormat { get; set; } = new[] {
        OfficeTableExportFormat.Csv, OfficeTableExportFormat.Markdown, OfficeTableExportFormat.Json
    };
    /// <summary>Write only visual metadata.</summary>
    [Parameter] public SwitchParameter NoPayload { get; set; }
    /// <summary>Write visual payloads without JSON metadata.</summary>
    [Parameter] public SwitchParameter NoMetadata { get; set; }
    /// <summary>Indent table and visual JSON.</summary>
    [Parameter] public SwitchParameter Indented { get; set; }
    /// <summary>Validate media payload hashes when the source supplies them.</summary>
    [Parameter] public SwitchParameter ValidatePayloadHash { get; set; }
    /// <summary>Leave existing output files unchanged and return skipped materialization results.</summary>
    [Parameter] public SwitchParameter NoClobber { get; set; }
    /// <summary>Worksheet name, one-based position or sheet object.</summary>
    [Parameter] [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }
    /// <summary>A1 range within the selected worksheet.</summary>
    [Parameter] public string? Range { get; set; }
    /// <summary>Maximum input bytes per source.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)]
    public long? MaximumInputBytes { get; set; }
    /// <summary>Caller-configured immutable Reader.</summary>
    [Parameter] public OfficeDocumentReader? Reader { get; set; }
    /// <summary>Typed common and format read policies, or common/format-keyed hashtables.</summary>
    [Parameter] [OfficeReadOptionsTransformation]
    public object[]? Options { get; set; }

    /// <inheritdoc />
    protected override Task BeginProcessingAsync() {
        _types = new HashSet<OfficeDocumentAssetType>(Type);
        _formats = new HashSet<OfficeTableExportFormat>(TableFormat);
        if (_types.Count == 0 || _types.Any(value => !Enum.IsDefined(typeof(OfficeDocumentAssetType), value)))
            throw new PSArgumentException("Choose Media, Table or Visual.");
        if (_formats.Count == 0 || _formats.Any(value => !Enum.IsDefined(typeof(OfficeTableExportFormat), value)))
            throw new PSArgumentException("Choose Csv, Markdown or Json table formats.");
        bool Bound(string name) => MyInvocation.BoundParameters.ContainsKey(name);
        if (!_types.Contains(OfficeDocumentAssetType.Media) && (Bound(nameof(Kind)) || Bound(nameof(MediaType)) || Bound(nameof(Extension)) || ValidatePayloadHash))
            throw new PSArgumentException("Kind, MediaType, Extension and ValidatePayloadHash apply to Media.");
        if (!_types.Contains(OfficeDocumentAssetType.Table) && Bound(nameof(TableFormat)))
            throw new PSArgumentException("TableFormat applies to Table.");
        if (!_types.Contains(OfficeDocumentAssetType.Visual) && (NoPayload || NoMetadata))
            throw new PSArgumentException("NoPayload and NoMetadata apply to Visual.");
        if (NoPayload && NoMetadata) throw new PSArgumentException("Choose visual payloads or metadata; both cannot be omitted.");
        _directory = OfficePath.ResolveOutput(SessionState, OutputDirectory, directory: true);
        _configuration = OfficeReadConfiguration.Bind(Options, Reader?.DefaultOptions);
        bool excel = Bound(nameof(Sheet)) || Bound(nameof(Range));
        if (Reader != null && (_configuration.HasFormatOptions || excel))
            throw new PSArgumentException("Configure format options when creating a caller-provided immutable Reader.");
        if (MaximumInputBytes.HasValue) _configuration.Common.MaxInputBytes = MaximumInputBytes.Value;
        if (excel) {
            var options = _configuration.Formats.Excel = OfficeReadConfiguration.Copy(_configuration.Formats.Excel);
            if (Sheet != null) {
                options.SheetNumber = Sheet.Position;
                options.SheetName = Sheet.Name ?? (Sheet.Sheet as OfficeIMO.Excel.ExcelSheet)?.Name;
            }
            if (Bound(nameof(Range))) options.A1Range = Range;
        }
        _reader = Reader ?? ReaderCommandUtilities.CreateReader(_configuration.Formats);
        _kinds = Normalize(Kind);
        _mediaTypes = Normalize(MediaType);
        _extensions = ReaderCommandUtilities.NormalizeSourceExtensions(Extension ?? Array.Empty<string>());
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override Task ProcessRecordAsync() {
        _paths.AddRange(ResolveInputs(allowDirectory: true));
        return Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async Task EndProcessingAsync() {
        try {
            await RunBatchAsync(_paths, async (path, token) => {
                if (!ShouldProcess(_directory, "Export document assets from " + path)) return null;
                OfficeDocumentReadResult document = await _reader.ReadDocumentAsync(path, _configuration.Common, token).ConfigureAwait(false);
                if (_types.Contains(OfficeDocumentAssetType.Media)) {
                    WriteObject(document.WriteAssetsToDirectory(_directory, new OfficeDocumentAssetMaterializationOptions {
                        Overwrite = !NoClobber, ValidatePayloadHash = ValidatePayloadHash,
                        Predicate = asset => Matches(_kinds, asset.Kind) && Matches(_mediaTypes, asset.MediaType) && Matches(_extensions, asset.Extension)
                    }, token), true);
                }
                if (_types.Contains(OfficeDocumentAssetType.Table)) {
                    WriteObject(_reader.ExportTables(document.EnumerateTables(), Indented, token)
                        .WriteTableExportsToDirectory(_directory, new ReaderTableExportMaterializationOptions {
                            Overwrite = !NoClobber,
                            IncludeCsv = _formats.Contains(OfficeTableExportFormat.Csv),
                            IncludeMarkdown = _formats.Contains(OfficeTableExportFormat.Markdown),
                            IncludeJson = _formats.Contains(OfficeTableExportFormat.Json)
                        }, token), true);
                }
                if (_types.Contains(OfficeDocumentAssetType.Visual)) {
                    WriteObject(_reader.ExportVisuals(document.EnumerateVisuals(), Indented, token)
                        .WriteVisualExportsToDirectory(_directory, new ReaderVisualExportMaterializationOptions {
                            Overwrite = !NoClobber, IncludePayload = !NoPayload, IncludeJson = !NoMetadata
                        }, token), true);
                }
                return null;
            }, "Export document assets", "asset-export-failed").ConfigureAwait(false);
        } finally { _configuration.Common.OpenPassword = null; }
    }

    private static string[] Normalize(string[]? values) => (values ?? Array.Empty<string>())
        .Where(static value => !string.IsNullOrWhiteSpace(value)).Select(static value => value.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();

    private static bool Matches(string[] filters, string? value) => filters.Length == 0 ||
        (!string.IsNullOrWhiteSpace(value) && filters.Contains(value, StringComparer.OrdinalIgnoreCase));
}
