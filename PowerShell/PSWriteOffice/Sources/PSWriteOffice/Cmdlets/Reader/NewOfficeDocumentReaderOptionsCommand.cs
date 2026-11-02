using System.Management.Automation;
using System.Text;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates common read, detection, package and folder policies for document readers.</summary>
/// <example><summary>Read text without derived Markdown or assets.</summary>
/// <code>Read-OfficeDocument ./Report.docx -As Text -Options (New-OfficeDocumentReaderOptions -Projection TextAndMetadata)</code></example>
/// <example><summary>Reuse a configured reader for a folder crawl.</summary>
/// <code>$options = New-OfficeDocumentReaderOptions -DetectionMode PreferContent -MaximumTraversalEntries 50000
/// $reader = New-OfficeDocumentReader -Options $options
/// Read-OfficeDocument ./Reports -Recurse -Reader $reader</code></example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentReaderOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(ReaderOptions))]
public sealed class NewOfficeDocumentReaderOptionsCommand : PSCmdlet {
    /// <summary>Derived representations to retain.</summary>
    [Parameter] public ReaderProjection Projection { get; set; } = ReaderProjection.Full;
    /// <summary>How extension and content evidence select the format handler.</summary>
    [Parameter] public ReaderDetectionMode DetectionMode { get; set; } = ReaderDetectionMode.ContentWhenUnknown;
    /// <summary>Maximum prefix bytes inspected during content detection.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumDetectionProbeBytes { get; set; }
    /// <summary>Maximum archive entries inspected during format detection.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumDetectionContainerEntries { get; set; }
    /// <summary>Maximum uncompressed bytes in one Office package part.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumPartUncompressedBytes { get; set; }
    /// <summary>Maximum combined uncompressed bytes in an Office package.</summary>
    [Parameter] [ValidateRange(1, long.MaxValue)] public long? MaximumTotalUncompressedBytes { get; set; }
    /// <summary>Maximum retained Office image assets; zero retains none.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaximumImageAssets { get; set; }
    /// <summary>Maximum bytes retained for one Office image asset.</summary>
    [Parameter] [ValidateRange(0, long.MaxValue)] public long? MaximumImageAssetBytes { get; set; }
    /// <summary>Maximum combined retained Office image payload bytes.</summary>
    [Parameter] [ValidateRange(0, long.MaxValue)] public long? MaximumTotalImageAssetBytes { get; set; }
    /// <summary>Maximum placements retained for one Office image relationship.</summary>
    [Parameter] [ValidateRange(0, int.MaxValue)] public int? MaximumImagePlacementsPerRelationship { get; set; }
    /// <summary>Maximum raw file-system entries inspected before filtering.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? MaximumTraversalEntries { get; set; }
    /// <summary>Follow links and junctions during folder traversal. Native format admission rules still apply.</summary>
    [Parameter] public SwitchParameter FollowSymlink { get; set; }
    /// <summary>Do not sort folder traversal results.</summary>
    [Parameter] public SwitchParameter NoDeterministicOrder { get; set; }
    /// <summary>Text encoding for sources without a byte-order mark.</summary>
    [Parameter] [OfficeEncodingTransformation] public Encoding? Encoding { get; set; }
    /// <summary>Reject invalid byte sequences instead of replacing them.</summary>
    [Parameter] public SwitchParameter FailOnInvalidTextBytes { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new ReaderOptions { Projection = Projection, DetectionMode = DetectionMode, ThrowOnInvalidTextBytes = FailOnInvalidTextBytes };
        if (Encoding != null) options.TextEncoding = (Encoding)Encoding.Clone();
        if (MaximumDetectionProbeBytes.HasValue) options.DetectionMaxProbeBytes = MaximumDetectionProbeBytes.Value;
        if (MaximumDetectionContainerEntries.HasValue) options.DetectionMaxContainerEntries = MaximumDetectionContainerEntries.Value;
        if (MaximumPartUncompressedBytes.HasValue) options.OpenXmlMaxPartUncompressedBytes = MaximumPartUncompressedBytes;
        if (MaximumTotalUncompressedBytes.HasValue) options.OpenXmlMaxTotalUncompressedBytes = MaximumTotalUncompressedBytes;
        if (MaximumImageAssets.HasValue) options.MaxOpenXmlImageAssets = MaximumImageAssets;
        if (MaximumImageAssetBytes.HasValue) options.MaxOpenXmlImageAssetBytes = MaximumImageAssetBytes;
        if (MaximumTotalImageAssetBytes.HasValue) options.MaxOpenXmlImageTotalAssetBytes = MaximumTotalImageAssetBytes;
        if (MaximumImagePlacementsPerRelationship.HasValue) options.MaxOpenXmlImagePlacementsPerRelationship = MaximumImagePlacementsPerRelationship;
        if (MaximumTraversalEntries.HasValue || MyInvocation.BoundParameters.ContainsKey(nameof(FollowSymlink)) || MyInvocation.BoundParameters.ContainsKey(nameof(NoDeterministicOrder))) {
            options.FolderOptions = new ReaderFolderOptions { Recurse = false, SkipReparsePoints = !FollowSymlink, DeterministicOrder = !NoDeterministicOrder };
            if (MaximumTraversalEntries.HasValue) options.FolderOptions.MaxTraversalEntries = MaximumTraversalEntries.Value;
        }
        WriteObject(options);
    }
}
