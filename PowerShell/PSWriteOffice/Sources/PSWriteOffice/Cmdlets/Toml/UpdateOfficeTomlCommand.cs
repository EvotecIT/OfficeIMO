using System.Management.Automation;
using OfficeIMO.Toml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Toml;

namespace PSWriteOffice.Cmdlets.Toml;

/// <summary>Applies a JSON Patch (RFC 6902) or JSON Merge Patch (RFC 7386) to a live TOML document.</summary>
/// <para>-JsonPatch takes a list of add, remove, replace, move, copy and test operations whose paths are JSON Pointers
/// such as /project/version. -MergePatch takes a partial document whose keys overwrite the document's, where $null
/// removes a key. Either can be JSON text or PowerShell hashtables, arrays and objects; dates in PowerShell values are
/// written as text.</para>
/// <para>Only the patched lines change; comments and layout elsewhere stay. New objects become [header] tables and
/// arrays of objects [[header]] arrays of tables. TOML has no null, so a null value in a JSON Patch fails. A patch is
/// all or nothing: when a test fails, a path is missing or an edit cannot be made, a non-terminating toml-patch-failed
/// error is written and the document is unchanged.</para>
/// <example><summary>Bump a version only if it is the expected one.</summary>
/// <code>Open-OfficeToml -Path ./pyproject.toml {
///     Update-OfficeToml -JsonPatch @(
///         @{ op = 'test'; path = '/project/version'; value = '1.4.0' }
///         @{ op = 'replace'; path = '/project/version'; value = '1.5.0' }
///         @{ op = 'add'; path = '/project/dependencies/-'; value = 'click' }
///     )
/// }</code></example>
/// <example><summary>Change a setting and remove another with a merge patch.</summary>
/// <code>Update-OfficeToml -Document $document -MergePatch @{ tool = @{ ruff = @{ 'line-length' = 100; exclude = $null } } } -PassThru</code></example>
[Cmdlet(VerbsData.Update, "OfficeToml", DefaultParameterSetName = JsonPatchSet)]
[OutputType(typeof(TomlPatchReport))]
public sealed class UpdateOfficeTomlCommand : OfficeDocumentCmdlet<TomlDocument> {
    private const string JsonPatchSet = "JsonPatch";
    private const string MergePatchSet = "MergePatch";

    /// <summary>The JSON Patch operations: JSON text, or an array of hashtables or objects.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = JsonPatchSet), ValidateNotNull]
    public object JsonPatch { get; set; } = null!;
    /// <summary>The merge patch: JSON text, or a hashtable or object; $null values remove keys.</summary>
    [Parameter(Mandatory = true, ParameterSetName = MergePatchSet), ValidateNotNull]
    public object MergePatch { get; set; } = null!;
    /// <summary>Return the report of added, replaced and removed paths.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TomlDocument document = GetDocument();
        bool merge = ParameterSetName == MergePatchSet;
        object input = merge ? MergePatch : JsonPatch;
        TomlPatchReport report;
        try {
            object patch = input is PSObject wrapped ? wrapped.BaseObject : input;
            string json;
            if (patch is string text) {
                json = text;
            } else {
                json = TomlPowerShellValue.ToJson(input, out bool truncated);
                if (truncated) WriteWarning("The patch is truncated: objects nested deeper than " + TomlPowerShellValue.MaxDepth + " levels are written as text.");
            }
            report = merge ? document.ApplyMergePatch(json) : document.ApplyJsonPatch(json);
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) {
            WriteError(OfficeTomlDocuments.Error(error, input));
            return;
        }
        if (PassThru) WriteObject(report);
    }
}
