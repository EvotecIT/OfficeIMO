using PSWriteOffice.Services;
using System.Management.Automation;
using System.Security;
using OfficeIMO.OpenDocument;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OpenDocument;

/// <summary>Loads a native ODT, ODS, ODP, or ODG document.</summary>
[Cmdlet(VerbsCommon.Open, "OfficeOpenDocument", DefaultParameterSetName = "Path", SupportsShouldProcess = true)]
[OutputType(typeof(OdfDocument), typeof(OdtDocument), typeof(OdsDocument), typeof(OdpPresentation), typeof(OdgDocument))]
public sealed class OpenOfficeOpenDocumentCommand : OfficeLifecycleCmdlet<OdfDocument>
{
    /// <summary>An optional authoring block; omission returns the live document.</summary>
    [Parameter(ParameterSetName = "Path", Position = 1)]
    [Parameter(ParameterSetName = "LiteralPath", Position = 1)]
    [ValidateNotNull]
    public override ScriptBlock? Content { get => base.Content; set => base.Content = value; }

    /// <summary>Emit the FileInfo after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter PassThru { get => base.PassThru; set => base.PassThru = value; }

    /// <summary>Open the saved document after a completed authoring block.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "LiteralPath")]
    public override SwitchParameter Show { get => base.Show; set => base.Show = value; }

    /// <summary>Path to an ODT, ODS, ODP, or ODG file.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;
    /// <summary>Exact source path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Password used to decrypt an encrypted OpenDocument package.</summary>
    [Parameter]
    [OfficePasswordTransformation]
    public SecureString? Password { get; set; }

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
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath"))
            RunDocument(path, false, false, source => OdfDocument.Load(source!, BuildOptions()),
                (document, destination, allowLoss) => {
                    var result = document.Save(destination);
                    if (!allowLoss) result.Report.ThrowIfLoss();
                    return null;
                });
    }
    /// <inheritdoc />
    protected override void InvokeContent(OdfDocument document) {
        using (PSWriteOffice.Services.OpenDocument.OpenDocumentDslContext.Enter(document)) Content!.InvokeReturnAsIs();
    }

    private OdfLoadOptions BuildOptions() {
        var options = new OdfLoadOptions();
        if (Password != null) options.Password = Password.ToPlainText();
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
