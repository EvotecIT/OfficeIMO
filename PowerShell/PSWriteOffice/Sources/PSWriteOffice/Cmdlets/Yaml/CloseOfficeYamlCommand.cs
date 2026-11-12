using System.Management.Automation;
using OfficeIMO.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Completes a live YAML document's lifetime without saving.</summary>
/// <example><summary>Discard pending changes.</summary>
/// <code>Close-OfficeYaml -Document $document</code></example>
/// <example><summary>Read a value and release the document.</summary>
/// <code>$document = Open-OfficeYaml -Path ./values.yaml
/// (Get-OfficeYamlNode -Document $document -Select 'image.tag').Value
/// Close-OfficeYaml -Document $document</code></example>
[Cmdlet(VerbsCommon.Close, "OfficeYaml", DefaultParameterSetName = "Document")]
public sealed class CloseOfficeYamlCommand : PSCmdlet {
    /// <summary>YAML document to close.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = "Document"), ValidateNotNull]
    public YamlDocument Document { get; set; } = null!;
    /// <inheritdoc />
    protected override void ProcessRecord() => Document.Dispose();
}
