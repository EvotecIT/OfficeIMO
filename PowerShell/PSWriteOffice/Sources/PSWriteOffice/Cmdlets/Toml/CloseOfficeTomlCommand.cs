using System.Management.Automation;
using OfficeIMO.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Completes a live TOML document's lifetime without saving.</summary>
/// <example><summary>Discard pending changes.</summary>
/// <code>Close-OfficeToml -Document $document</code></example>
/// <example><summary>Read a value and release the document.</summary>
/// <code>$document = Open-OfficeToml -Path ./Cargo.toml
/// (Get-OfficeTomlNode -Document $document -Select 'package.version').Value
/// Close-OfficeToml -Document $document</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeToml", DefaultParameterSetName = "Document")]
public sealed class CloseOfficeTomlCommand : PSCmdlet {
    /// <summary>TOML document to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public TomlDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
