using System.Management.Automation;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Loads a native ODT, ODS, ODP, or ODG document.</summary>
[Cmdlet(VerbsCommon.Get, "OfficeOpenDocument")]
[OutputType(typeof(OdfDocument), typeof(OdtDocument), typeof(OdsDocument), typeof(OdpPresentation), typeof(OdgDocument))]
public sealed class GetOfficeOpenDocumentCommand : PSCmdlet
{
    /// <summary>Path to an ODT, ODS, ODP, or ODG file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional bounded package and XML settings.</summary>
    [Parameter]
    public OdfLoadOptions? Options { get; set; }

    /// <summary>Password used to decrypt an encrypted OpenDocument package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Maximum source package size in bytes.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long? MaxPackageBytes { get; set; }

    /// <summary>Maximum number of ZIP entries.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxEntries { get; set; }

    /// <summary>Maximum uncompressed size of one package entry.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long? MaxEntryUncompressedBytes { get; set; }

    /// <summary>Maximum aggregate uncompressed package size.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long? MaxTotalUncompressedBytes { get; set; }

    /// <summary>Maximum aggregate PBKDF2 iterations across encrypted entries.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long? MaxTotalKdfIterations { get; set; }

    /// <summary>Maximum declared expansion ratio for a compressed entry.</summary>
    [Parameter]
    [ValidateRange(double.Epsilon, double.MaxValue)]
    public double? MaxCompressionRatio { get; set; }

    /// <summary>Maximum archive path depth.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxDepth { get; set; }

    /// <summary>Maximum characters allowed in one parsed XML part.</summary>
    [Parameter]
    [ValidateRange(1L, long.MaxValue)]
    public long? MaxXmlCharacters { get; set; }

    /// <summary>Maximum element nesting depth in one parsed XML part.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxXmlDepth { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(OdfDocument.Load(
        SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path), BuildOptions()));

    private OdfLoadOptions BuildOptions() {
        var options = Options == null
            ? new OdfLoadOptions()
            : new OdfLoadOptions {
                Password = Options.Password,
                MaxPackageBytes = Options.MaxPackageBytes,
                MaxEntries = Options.MaxEntries,
                MaxEntryUncompressedBytes = Options.MaxEntryUncompressedBytes,
                MaxTotalUncompressedBytes = Options.MaxTotalUncompressedBytes,
                MaxTotalKdfIterations = Options.MaxTotalKdfIterations,
                MaxCompressionRatio = Options.MaxCompressionRatio,
                MaxDepth = Options.MaxDepth,
                MaxXmlCharacters = Options.MaxXmlCharacters,
                MaxXmlDepth = Options.MaxXmlDepth
            };
        if (Password != null) options.Password = Password;
        if (MaxPackageBytes.HasValue) options.MaxPackageBytes = MaxPackageBytes.Value;
        if (MaxEntries.HasValue) options.MaxEntries = MaxEntries.Value;
        if (MaxEntryUncompressedBytes.HasValue) options.MaxEntryUncompressedBytes = MaxEntryUncompressedBytes.Value;
        if (MaxTotalUncompressedBytes.HasValue) options.MaxTotalUncompressedBytes = MaxTotalUncompressedBytes.Value;
        if (MaxTotalKdfIterations.HasValue) options.MaxTotalKdfIterations = MaxTotalKdfIterations.Value;
        if (MaxCompressionRatio.HasValue) options.MaxCompressionRatio = MaxCompressionRatio.Value;
        if (MaxDepth.HasValue) options.MaxDepth = MaxDepth.Value;
        if (MaxXmlCharacters.HasValue) options.MaxXmlCharacters = MaxXmlCharacters.Value;
        if (MaxXmlDepth.HasValue) options.MaxXmlDepth = MaxXmlDepth.Value;
        return options;
    }
}
