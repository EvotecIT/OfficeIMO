using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Reads TOML files as PowerShell objects; the file form of ConvertFrom-OfficeToml.</summary>
/// <para>Output follows ConvertFrom-OfficeToml: tables become PSCustomObjects in file order, or ordered hashtables with
/// AsHashtable. A file that fails to parse writes a non-terminating error with the engine's code, line and column, and
/// the next file is read.</para>
/// <example><summary>Read a tool's settings.</summary>
/// <code>$ruff = (Import-OfficeToml -Path ./pyproject.toml).tool.ruff
/// $ruff.'line-length'</code></example>
/// <example><summary>Read the manifest of every crate under a folder.</summary>
/// <code>Get-ChildItem ./crates -Filter Cargo.toml -Recurse | Import-OfficeToml | ForEach-Object { $_.package.name }</code></example>
[Cmdlet(VerbsData.Import, "OfficeToml", DefaultParameterSetName = "Path")]
[OutputType(typeof(PSObject), typeof(System.Collections.Specialized.OrderedDictionary))]
public sealed class ImportOfficeTomlCommand : OfficeCmdlet {
    /// <summary>TOML files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal TOML files, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>Return tables as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }
    /// <summary>Deepest nesting of tables and arrays to accept; sets the reader's MaxDepth.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Depth { get; set; }
    /// <summary>Caps every limit at the engine's untrusted defaults, for files from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlLoadOptions options = OfficeTomlDocuments.LoadOptions(Untrusted.IsPresent, MyInvocation.BoundParameters.ContainsKey(nameof(Depth)) ? Depth : null);
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            try {
                using TomlDocument document = TomlDocument.Load(source, options);
                OfficeTomlDocuments.WriteData(this, document, AsHashtable.IsPresent);
            } catch (PipelineStoppedException) { throw; }
            catch (Exception error) when (error is TomlException or InvalidDataException or IOException or UnauthorizedAccessException) {
                WriteError(OfficeTomlDocuments.Error(error, source));
            }
        }
    }
}
