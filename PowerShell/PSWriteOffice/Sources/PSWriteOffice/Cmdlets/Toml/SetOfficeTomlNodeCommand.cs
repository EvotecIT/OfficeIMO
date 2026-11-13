using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Sets the value at a path in a live TOML document, adding missing tables on the way.</summary>
/// <para>Setting a value over a value changes only its text, and an equal value keeps its original spelling, so saving
/// changes only what really changed. A new key is written after its table's last key in the indentation and spacing of
/// its neighbors; a missing table gets a [header] section. Hashtables and objects become inline tables, and arrays
/// become arrays. TOML has no null, so $null is an error and nulls inside a value are omitted with a warning. An array
/// index must already exist.</para>
/// <example><summary>Bump a version inside an editing block.</summary>
/// <code>Open-OfficeToml -Path ./Cargo.toml {
///     Set-OfficeTomlNode -Select 'package.version' -Value '0.4.0'
/// }</code></example>
/// <example><summary>Add an inline table and replace an array item.</summary>
/// <code>Set-OfficeTomlNode -Document $document -Select 'dependencies.serde' -Value ([ordered]@{ version = '1'; features = @('derive') })
/// Set-OfficeTomlNode -Document $document -Select 'project.dependencies[0]' -Value 'httpx>=0.27'</code></example>
[Cmdlet(VerbsCommon.Set, "OfficeTomlNode", DefaultParameterSetName = "Document")]
[OutputType(typeof(PSObject))]
public sealed class SetOfficeTomlNodeCommand : OfficeDocumentCmdlet<TomlDocument> {
    /// <summary>Path of the key or array item to set.</summary>
    [Parameter(Mandatory = true, Position = 0), ValidateNotNullOrEmpty]
    public string Select { get; set; } = string.Empty;
    /// <summary>The new value: a string, number, Boolean, date, hashtable, array or object.</summary>
    [Parameter(Mandatory = true, Position = 1), AllowNull, AllowEmptyString, AllowEmptyCollection]
    public object? Value { get; set; }
    /// <summary>Return the node that is now at the path.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlDocument document = GetDocument();
        TomlNode node;
        try {
            object value = TomlPowerShellValue.ToValue(Value, out int nulls, out bool truncated);
            node = document.Set(Select, value);
            if (truncated) WriteWarning("The value is truncated: objects nested deeper than " + TomlPowerShellValue.MaxDepth + " levels are written as text.");
            if (nulls > 0) WriteWarning("TOML has no null, so " + nulls + " null value" + (nulls == 1 ? " was" : "s were") + " omitted.");
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) {
            WriteError(OfficeTomlDocuments.Error(error, Select));
            return;
        }
        if (PassThru) WriteObject(OfficeTomlDocuments.DescribeNode(node, Select, asHashtable: false));
    }
}
