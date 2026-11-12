using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Reads YAML files as PowerShell objects; the file form of ConvertFrom-OfficeYaml.</summary>
/// <para>The encoding (UTF-8, UTF-16 or UTF-32) is detected from the file. Output follows ConvertFrom-OfficeYaml: mappings
/// become PSCustomObjects in file order, or ordered hashtables with AsHashtable. A file that fails to parse writes a
/// non-terminating error with the engine's code, line and column, and the next file is read.</para>
/// <example><summary>Read a settings file.</summary>
/// <code>$settings = Import-OfficeYaml -Path ./settings.yaml
/// $settings.logging.level</code></example>
/// <example><summary>List the kinds in every manifest under a folder.</summary>
/// <code>Get-ChildItem ./deploy -Filter *.yaml | Import-OfficeYaml -AllDocuments | Group-Object kind</code></example>
[Cmdlet(VerbsData.Import, "OfficeYaml", DefaultParameterSetName = "Path")]
[OutputType(typeof(PSObject), typeof(System.Collections.Specialized.OrderedDictionary), typeof(object[]))]
public sealed class ImportOfficeYamlCommand : OfficeCmdlet {
    /// <summary>YAML files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal YAML files, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Return mappings as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }
    /// <summary>Deepest nesting of mappings and sequences to accept; sets the reader's MaxDepth.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Depth { get; set; }
    /// <summary>Write a root sequence as one array instead of enumerating its items.</summary>
    [Parameter]
    public SwitchParameter NoEnumerate { get; set; }
    /// <summary>Read every document of a multi-document file; without it, more than one document is an error.</summary>
    [Parameter]
    public SwitchParameter AllDocuments { get; set; }
    /// <summary>Schema for scalar types: Core (the default), Json, Yaml11, or Failsafe to keep every scalar a string.</summary>
    [Parameter]
    public YamlSchema? Schema { get; set; }
    /// <summary>Applies the YAML specification's exact rules instead of accepting common real-world leniencies.</summary>
    [Parameter]
    public SwitchParameter Strict { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults, for files from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlLoadOptions options = OfficeYamlDocuments.BuildLoadOptions(null, Untrusted.IsPresent, BoundSwitch(nameof(Strict)), Schema,
            MyInvocation.BoundParameters.ContainsKey(nameof(Depth)) ? Depth : null);
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            try {
                using YamlDocument document = YamlDocument.Load(source, options);
                OfficeYamlDocuments.WriteData(this, document, AllDocuments.IsPresent, AsHashtable.IsPresent, NoEnumerate.IsPresent);
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) when (error is YamlException or InvalidDataException or IOException or UnauthorizedAccessException) {
                WriteError(OfficeYamlDocuments.Error(error, source));
            }
        }
    }
}
