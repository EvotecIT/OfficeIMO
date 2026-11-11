using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Sets the value at a path in a live YAML document, creating missing mappings on the way.</summary>
/// <para>Setting a scalar over a scalar keeps its comments, anchor and tag, and an equal value keeps its original
/// spelling, so saving changes only the text of what really changed. Hashtables, arrays and objects become mappings
/// and sequences in the layout of their new neighbors. An index equal to a sequence's length appends.</para>
/// <example><summary>Bump an image tag inside an editing block.</summary>
/// <code>Open-OfficeYaml -Path ./values.yaml {
///     Set-OfficeYamlNode -Select 'image.tag' -Value '1.27.2'
/// }</code></example>
/// <example><summary>Add a mapping and append a sequence item.</summary>
/// <code>Set-OfficeYamlNode -Document $document -Select 'resources.limits' -Value ([ordered]@{ cpu = '500m'; memory = '256Mi' })
/// Set-OfficeYamlNode -Document $document -Select 'ports[2]' -Value 8443</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeYamlNode", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject))]
public sealed class SetOfficeYamlNodeCommand : OfficeDocumentCmdlet<YamlDocument> {
    /// <summary>Path of the node to set; $ sets the root of the first document.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNull]
    public string Select { get; set; } = string.Empty;
    /// <summary>The new value: a scalar, hashtable, array, object or YamlNode. $null writes a YAML null.</summary>
    [Parameter(Mandatory = true, Position = 1), AllowNull, AllowEmptyString, AllowEmptyCollection]
    public object? Value { get; set; }
    /// <summary>Scalar style for a string value, such as DoubleQuoted or Literal; omission picks a safe style.</summary>
    [Parameter]
    public YamlStyle? Style { get; set; }
    /// <summary>Return the node that is now at the path.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlDocument document = GetDocument();
        YamlNode node;
        try {
            node = document.Set(Select, YamlPowerShellValue.ToNode(Value, Style, out bool truncated));
            if (truncated) WriteWarning("The value is truncated: objects nested deeper than " + YamlPowerShellValue.MaxDepth + " levels are written as text.");
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) {
            WriteError(OfficeYamlDocuments.Error(error, Select));
            return;
        }
        if (PassThru) WriteObject(OfficeYamlDocuments.DescribeNode(node, Select, asHashtable: false));
    }
}
