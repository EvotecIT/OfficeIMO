using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Removes mapping entries or sequence items from a live YAML document by path.</summary>
/// <para>An entry goes with its own comment lines and line; neighboring entries and section comments stay. A path with
/// nothing at it writes a non-terminating ObjectNotFound error.</para>
/// <example><summary>Drop a setting inside an editing block.</summary>
/// <code>Open-OfficeYaml -Path ./values.yaml {
///     Remove-OfficeYamlNode -Select 'ingress.annotations["nginx.ingress.kubernetes.io/rewrite-target"]'
/// }</code></example>
/// <example><summary>Remove a sequence item if it is there.</summary>
/// <code>Remove-OfficeYamlNode -Document $document -Select 'spec.ports[1]' -ErrorAction SilentlyContinue</code></example>
[Cmdlet(VerbsCommon.Remove, "OfficeYamlNode", DefaultParameterSetName = "Document")]
public sealed class RemoveOfficeYamlNodeCommand : OfficeDocumentCmdlet<YamlDocument> {
    /// <summary>Paths of the entries or items to remove.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string[] Select { get; set; } = Array.Empty<string>();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlDocument document = GetDocument();
        foreach (string path in Select) {
            try {
                if (!document.Remove(path)) throw new ItemNotFoundException("No YAML node is at '" + path + "'.");
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) { WriteError(OfficeYamlDocuments.Error(error, path)); }
        }
    }
}
