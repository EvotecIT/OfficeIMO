using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Gets nodes of a live YAML document by path, as Path, Value, Node, Line and Column objects.</summary>
/// <para>Select uses the engine's path grammar: dotted keys, [n] for 0-based sequence items, ["key.with.dots"] for keys
/// that are not plain names, and a leading $[n] for the 0-based document in a multi-document stream. Value holds plain
/// PowerShell data as ConvertFrom-OfficeYaml returns it; Node is the live node. Line and Column are 1-based. Without
/// Select, each document's root is returned.</para>
/// <example><summary>Read one value.</summary>
/// <code>(Get-OfficeYamlNode -Document $document -Select 'spec.template.spec.containers[0].image').Value</code></example>
/// <example><summary>Find where settings live in the file.</summary>
/// <code>Get-OfficeYamlNode -Document $document -Select 'spec.replicas', 'metadata.labels["app.kubernetes.io/name"]' |
///     Select-Object -Property Path, Value, Line, Column</code></example>
[Cmdlet(VerbsCommon.Get, "OfficeYamlNode", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject))]
public sealed class GetOfficeYamlNodeCommand : OfficeDocumentCmdlet<YamlDocument> {
    /// <summary>Paths of the nodes to get; omission returns each document's root.</summary>
    [Parameter(Position = 0), ValidateNotNullOrEmpty]
    public string[]? Select { get; set; }
    /// <summary>Return mappings in Value as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlDocument document = GetDocument();
        if (Select == null) {
            IReadOnlyList<YamlDocumentPart> parts = document.Documents;
            for (int i = 0; i < parts.Count; i++) {
                string path = parts.Count == 1 ? "$" : "$[" + i.ToString(CultureInfo.InvariantCulture) + "]";
                WriteObject(OfficeYamlDocuments.DescribeNode(parts[i].Root, path, AsHashtable.IsPresent));
            }
            return;
        }
        foreach (string path in Select) {
            try {
                WriteObject(OfficeYamlDocuments.DescribeNode(OfficeYamlDocuments.Select(document, path), path, AsHashtable.IsPresent));
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) { WriteError(OfficeYamlDocuments.Error(error, path)); }
        }
    }
}
