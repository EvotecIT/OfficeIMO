using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Converts PowerShell objects to YAML text, as ConvertTo-Json does for JSON.</summary>
/// <para>Hashtables, ordered dictionaries and object properties become mappings, collections become sequences and
/// enums their names. One piped object is written as itself and several as a sequence. Objects nested deeper than
/// Depth (default 2) are written as their string form with a warning, as ConvertTo-Json does. Strings that YAML 1.1
/// readers would take as another type (yes, on, 012, 1e3) are quoted.</para>
/// <example><summary>Turn a hashtable into a configuration snippet.</summary>
/// <code>[ordered]@{ name = 'web'; replicas = 3; ports = 80, 443 } | ConvertTo-OfficeYaml</code></example>
/// <example><summary>Write nested service data with four-space indentation.</summary>
/// <code>Get-Service | Select-Object -First 3 Name, Status, DependentServices |
///     ConvertTo-OfficeYaml -Depth 3 -WriteOptions @{ MappingIndent = 4; SequenceIndent = 4; SequenceDashOffset = 2 }</code></example>
[Cmdlet(VerbsData.ConvertTo, "OfficeYaml", DefaultParameterSetName = "InputObject")]
[OutputType(typeof(string))]
public sealed class ConvertToOfficeYamlCommand : OfficeCmdlet {
    private readonly OfficeYamlWriterInput _input = new();

    /// <summary>Objects to convert; an argument is written exactly as given, piped objects as one value or a sequence.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowNull, AllowEmptyCollection]
    public object? InputObject { get; set; }
    /// <summary>How many levels of nested objects to write; deeper values become their string form.</summary>
    [Parameter, ValidateRange(0, YamlPowerShellValue.MaxDepth)]
    public int Depth { get; set; } = 2;
    /// <summary>Indentation and line-break settings, as a YamlWriterOptions object or hashtable.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(YamlWriterOptions)), ValidateNotNull]
    public YamlWriterOptions? WriteOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => _input.Add(this, InputObject);

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (!_input.HasInput) return;
        try {
            WriteObject(_input.Serialize(this, Depth, WriteOptions));
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) when (error is YamlException or ArgumentException or InvalidOperationException or NotSupportedException) {
            ThrowTerminatingError(OfficeYamlDocuments.Error(error, InputObject));
        }
    }
}
