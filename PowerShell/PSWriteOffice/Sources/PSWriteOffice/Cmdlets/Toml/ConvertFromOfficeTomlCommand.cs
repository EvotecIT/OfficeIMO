using System.Management.Automation;
using System.Text;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Converts TOML text to PowerShell objects, as ConvertFrom-Json does for JSON.</summary>
/// <para>Tables become PSCustomObjects with properties in file order (ordered hashtables with AsHashtable), arrays and
/// arrays of tables become arrays, integers Int64, floats Double, and date-times DateTimeOffset (with an offset),
/// DateTime (local date-time or date) or TimeSpan (local time). Piped lines are joined before parsing, so Get-Content
/// output works directly. Depth sets the deepest nesting accepted.</para>
/// <example><summary>Read a project manifest's values.</summary>
/// <code>$project = Get-Content -Path ./pyproject.toml -Raw | ConvertFrom-OfficeToml
/// $project.project.dependencies</code></example>
/// <example><summary>List the packages of a Cargo lock file as hashtables.</summary>
/// <code>(Get-Content ./Cargo.lock | ConvertFrom-OfficeToml -AsHashtable).package | ForEach-Object { '{0} {1}' -f $_.name, $_.version }</code></example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeToml", DefaultParameterSetName = "InputObject")]
[OutputType(typeof(PSObject), typeof(System.Collections.Specialized.OrderedDictionary))]
public sealed class ConvertFromOfficeTomlCommand : OfficeCmdlet {
    private readonly StringBuilder _text = new();
    private bool _hasInput;

    /// <summary>TOML text; piped strings are joined with line breaks before parsing.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "InputObject"), AllowEmptyString]
    public string InputObject { get; set; } = string.Empty;
    /// <summary>Return tables as ordered hashtables instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }
    /// <summary>Deepest nesting of tables and arrays to accept; sets the reader's MaxDepth.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Depth { get; set; }
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
        TomlLoadOptions options = OfficeTomlDocuments.LoadOptions(Untrusted.IsPresent, MyInvocation.BoundParameters.ContainsKey(nameof(Depth)) ? Depth : null);
        try {
            using TomlDocument document = TomlDocument.Parse(_text.ToString(), options);
            OfficeTomlDocuments.WriteData(this, document, AsHashtable.IsPresent);
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) when (error is TomlException or InvalidDataException) {
            ThrowTerminatingError(OfficeTomlDocuments.Error(error, null));
        }
    }
}
