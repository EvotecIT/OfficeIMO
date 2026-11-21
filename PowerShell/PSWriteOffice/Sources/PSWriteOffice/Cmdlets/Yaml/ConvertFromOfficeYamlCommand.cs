using System.Management.Automation;
using System.Text;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Converts YAML text to PowerShell objects, as ConvertFrom-Json does for JSON.</summary>
/// <para>Mappings become PSCustomObjects with properties in file order (ordered hashtables with AsHashtable), sequences
/// become arrays, and scalars resolve through the Core schema to Int64, BigInteger, Double, Boolean, $null or String.
/// Schema Failsafe keeps every scalar a string; Schema Yaml11 reads yes/no/on/off as Booleans. Aliases and merge keys
/// are resolved. Piped lines are joined before parsing, so Get-Content output works directly. Depth sets the deepest
/// nesting accepted (the engine default is 256).</para>
/// <example><summary>Read a configuration file's values.</summary>
/// <code>$config = Get-Content -Path ./settings.yaml -Raw | ConvertFrom-OfficeYaml
/// $config.database.port</code></example>
/// <example><summary>Read every manifest in a multi-document stream as hashtables.</summary>
/// <code>Get-Content ./manifests.yaml | ConvertFrom-OfficeYaml -AllDocuments -AsHashtable |
///     ForEach-Object { '{0}/{1}' -f $_.kind, $_.metadata.name }</code></example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeYaml", DefaultParameterSetName = "InputObject")]
[OutputType(typeof(PSObject), typeof(System.Collections.Specialized.OrderedDictionary), typeof(object[]))]
public sealed class ConvertFromOfficeYamlCommand : OfficeCmdlet {
    private readonly StringBuilder _text = new();
    private bool _hasInput;

    /// <summary>YAML text; piped strings are joined with line breaks before parsing.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowEmptyString]
    public string InputObject { get; set; } = string.Empty;
    /// <summary>Return mappings as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }
    /// <summary>Deepest nesting of mappings and sequences to accept; sets the reader's MaxDepth.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Depth { get; set; }
    /// <summary>Write a root sequence as one array instead of enumerating its items.</summary>
    [Parameter]
    public SwitchParameter NoEnumerate { get; set; }
    /// <summary>Read every document of a multi-document stream; without it, more than one document is an error.</summary>
    [Parameter]
    public SwitchParameter AllDocuments { get; set; }
    /// <summary>Schema for scalar types: Core (the default), Json, Yaml11, or Failsafe to keep every scalar a string.</summary>
    [Parameter]
    public YamlSchema? Schema { get; set; }
    /// <summary>Applies the YAML specification's exact rules instead of accepting common real-world leniencies.</summary>
    [Parameter]
    public SwitchParameter Strict { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults, for text from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (_hasInput) _text.Append('\n');
        _text.Append(InputObject);
        _hasInput = true;
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        YamlLoadOptions options = OfficeYamlDocuments.BuildLoadOptions(null, Untrusted.IsPresent, BoundSwitch(nameof(Strict)), Schema,
            MyInvocation.BoundParameters.ContainsKey(nameof(Depth)) ? Depth : null);
        try {
            using YamlDocument document = YamlDocument.Parse(_text.ToString(), options);
            OfficeYamlDocuments.WriteData(this, document, AllDocuments.IsPresent, AsHashtable.IsPresent, NoEnumerate.IsPresent);
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) when (error is YamlException or InvalidDataException) {
            ThrowTerminatingError(OfficeYamlDocuments.Error(error, null));
        }
    }
}
