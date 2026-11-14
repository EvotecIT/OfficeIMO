using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Toml;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Tests whether TOML files or text load, without changing them.</summary>
/// <para>Returns $true or $false for each input, as Test-Json does. Detailed returns a result with the error's
/// diagnostic code, message, and 1-based line and column. Loading checks all of TOML 1.1, including duplicate keys and
/// the rules for defining tables; Untrusted applies the engine's limits for files from outside.</para>
/// <example><summary>Gate a pipeline step on a valid file.</summary>
/// <code>if (-not (Test-OfficeToml -Path ./pyproject.toml)) { throw 'pyproject.toml is not valid TOML.' }</code></example>
/// <example><summary>Report where every broken file fails.</summary>
/// <code>Get-ChildItem ./crates -Recurse -Filter *.toml | Test-OfficeToml -Detailed |
///     Where-Object { -not $_.IsValid } |
///     Select-Object -Property Path, Line, Column, Code, Message</code></example>
[Cmdlet(VerbsDiagnostic.Test, "OfficeToml", DefaultParameterSetName = "Path")]
[OutputType(typeof(bool), typeof(OfficeTomlTestResult))]
public sealed class TestOfficeTomlCommand : OfficeCmdlet {
    /// <summary>TOML files with FileSystem wildcard expansion.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    [SupportsWildcards, ValidateNotNullOrEmpty]
    public string[] Path { get; set; } = Array.Empty<string>();
    /// <summary>Literal TOML files, including FileInfo.FullName and provider PSPath.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("FullName", "PSPath"), ValidateNotNullOrEmpty]
    public string[] LiteralPath { get; set; } = Array.Empty<string>();
    /// <summary>TOML text to test.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Text"), AllowEmptyString]
    public string Text { get; set; } = string.Empty;
    /// <summary>Caps every limit at the engine's untrusted defaults, for files from outside.</summary>
    [Parameter]
    public SwitchParameter Untrusted { get; set; }
    /// <summary>Return a result with the error code and position instead of a Boolean.</summary>
    [Parameter]
    public SwitchParameter Detailed { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlLoadOptions options = OfficeTomlDocuments.LoadOptions(Untrusted.IsPresent, null);
        if (ParameterSetName == "Text") {
            Write(null, () => TomlDocument.Parse(Text, options));
            return;
        }
        bool literal = ParameterSetName == "LiteralPath";
        foreach (string source in OfficePath.ResolveInputs(this, literal ? LiteralPath : Path, literal)) {
            Write(source, () => TomlDocument.Load(source, options));
        }
    }

    private void Write(string? path, Func<TomlDocument> load) {
        OfficeTomlTestResult result;
        try {
            using TomlDocument document = load();
            result = new OfficeTomlTestResult(path);
        } catch (TomlException error) {
            result = new OfficeTomlTestResult(path, error);
        }
        if (Detailed) WriteObject(result);
        else WriteObject(result.IsValid);
    }
}
