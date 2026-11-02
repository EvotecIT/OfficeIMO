using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Reader;
using OfficeIMO.Security;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Lists installed reading, conversion and protection capabilities from their canonical catalogs.</summary>
/// <example><summary>Find executable Word-to-PDF routes and their fidelity limits.</summary>
/// <code>Get-OfficeDocumentCapability -Conversion -From Word -To Pdf -Executable | Format-List *</code></example>
/// <example><summary>Inspect PDF protection support.</summary>
/// <code>Get-OfficeDocumentCapability -Protection -Format PDF</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeDocumentCapability", DefaultParameterSetName = "Read")]
[OutputType(typeof(ReaderHandlerCapability), typeof(ReaderCapabilityManifest), typeof(OfficeConversionCapability), typeof(OfficeProtectionCapability), typeof(string))]
public sealed partial class GetOfficeDocumentCapabilityCommand : PSCmdlet {
    /// <summary>Immutable caller-configured Reader registry.</summary>
    [Parameter(ParameterSetName = "Read")] public OfficeDocumentReader? Reader { get; set; }
    /// <summary>Return handlers supplied by OfficeIMO or registered by a caller.</summary>
    [Parameter(ParameterSetName = "Read")] public ReaderHandlerOrigin? Origin { get; set; }
    /// <summary>Return the Reader capability manifest instead of individual handlers.</summary>
    [Parameter(ParameterSetName = "Read")] public SwitchParameter Manifest { get; set; }
    /// <summary>List conversion routes and their fidelity and support evidence.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Conversion")] public SwitchParameter Conversion { get; set; }
    /// <summary>Source format family or exact extension, with or without a leading dot.</summary>
    [Parameter(ParameterSetName = "Conversion")] [ArgumentCompleter(typeof(OfficeFormatCompleter))]
    public string? From { get; set; }
    /// <summary>Destination format family or exact extension, with or without a leading dot.</summary>
    [Parameter(ParameterSetName = "Conversion")] [ArgumentCompleter(typeof(OfficeFormatCompleter))]
    public string? To { get; set; }
    /// <summary>Exact stable route or protection capability identifiers.</summary>
    [Parameter(ValueFromPipelineByPropertyName = true, ParameterSetName = "Conversion")]
    [Parameter(Position = 0, ParameterSetName = "Protection")]
    public string[]? Id { get; set; }
    /// <summary>Return only conversion routes implemented by this installed host.</summary>
    [Parameter(ParameterSetName = "Conversion")] public SwitchParameter Executable { get; set; }
    /// <summary>List password, encryption, signature and restriction support.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Protection")] public SwitchParameter Protection { get; set; }
    /// <summary>Protection format identifier or family, such as PDF or DOCX.</summary>
    [Parameter(ParameterSetName = "Protection")] public string? Format { get; set; }
    /// <summary>Protected-content mechanism.</summary>
    [Parameter(ParameterSetName = "Protection")] public OfficeProtectionKind? Kind { get; set; }
    /// <summary>Return protection rows with unsupported operations.</summary>
    [Parameter(ParameterSetName = "Protection")] public SwitchParameter IncompleteOnly { get; set; }
    /// <summary>Return the complete protection catalog as deterministic JSON, without row filters.</summary>
    [Parameter(ParameterSetName = "Protection")] public SwitchParameter AsJson { get; set; }
    /// <summary>Return the complete protection catalog as Markdown, without row filters.</summary>
    [Parameter(ParameterSetName = "Protection")] public SwitchParameter AsMarkdown { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == "Conversion") { WriteConversions(); return; }
        if (ParameterSetName == "Protection") { WriteProtection(); return; }
        var reader = Reader ?? ReaderCommandUtilities.Reader;
        var capabilities = reader.GetCapabilities().Where(capability => !Origin.HasValue || capability.Origin == Origin.Value).ToArray();
        if (Manifest) {
            var manifest = reader.GetCapabilityManifest(); manifest.Handlers = capabilities; WriteObject(manifest);
        } else WriteObject(capabilities, enumerateCollection: true);
    }
}
