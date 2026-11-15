using System.Management.Automation;
using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Yaml;

namespace PSWriteOffice.Cmdlets.Yaml;

/// <summary>Applies a JSON Patch (RFC 6902) or JSON Merge Patch (RFC 7386) to a live YAML document.</summary>
/// <para>-JsonPatch takes a list of add, remove, replace, move, copy and test operations whose paths are JSON Pointers
/// such as /spec/replicas. -MergePatch takes a partial document whose keys overwrite the document's, where $null removes
/// a key. Either can be JSON or YAML text, a YAML document, or PowerShell hashtables, arrays and objects.</para>
/// <para>Only the patched lines change; comments, quoting and layout elsewhere stay. A patch is all or nothing: when a
/// test fails, a path is missing or an edit cannot be made, a non-terminating yaml-patch-failed error is written and the
/// document is unchanged. Patches address the first document of a multi-document stream.</para>
/// <example><summary>Scale a deployment only if it still has the expected replica count.</summary>
/// <code>Open-OfficeYaml -Path ./deployment.yaml {
///     Update-OfficeYaml -JsonPatch @(
///         @{ op = 'test'; path = '/spec/replicas'; value = 2 }
///         @{ op = 'replace'; path = '/spec/replicas'; value = 3 }
///     )
/// }</code></example>
/// <example><summary>Apply a patch file and list what changed.</summary>
/// <code>Update-OfficeYaml -Document $document -JsonPatch (Get-Content ./scale.patch.json -Raw) -PassThru</code></example>
/// <example><summary>Override values and remove a key with a merge patch.</summary>
/// <code>Update-OfficeYaml -Document $document -MergePatch @{ image = @{ tag = '1.27.2' }; podAnnotations = $null }</code></example>
[Cmdlet(VerbsData.Update, "OfficeYaml", DefaultParameterSetName = JsonPatchSet)]
[OutputType(typeof(YamlPatchReport))]
public sealed class UpdateOfficeYamlCommand : OfficeDocumentCmdlet<YamlDocument> {
    private const string JsonPatchSet = "JsonPatch";
    private const string MergePatchSet = "MergePatch";

    /// <summary>The JSON Patch operations: JSON or YAML text, a YAML document, or an array of hashtables or objects.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = JsonPatchSet), ValidateNotNull]
    public object JsonPatch { get; set; } = null!;
    /// <summary>The merge patch: JSON or YAML text, a YAML document, or a hashtable or object; $null values remove keys.</summary>
    [Parameter(Mandatory = true, ParameterSetName = MergePatchSet), AllowNull]
    public object? MergePatch { get; set; }
    /// <summary>Return the report of added, replaced and removed paths.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        YamlDocument document = GetDocument();
        bool merge = ParameterSetName == MergePatchSet;
        object? input = merge ? MergePatch : JsonPatch;
        YamlPatchReport report;
        try {
            object? patch = input is PSObject wrapped ? wrapped.BaseObject : input;
            if (patch is string text) {
                report = merge ? document.ApplyMergePatch(text) : document.ApplyJsonPatch(text);
            } else if (patch is YamlDocument patchDocument) {
                report = merge ? document.ApplyMergePatch(patchDocument) : document.ApplyJsonPatch(patchDocument);
            } else {
                using YamlDocument built = YamlDocument.Create();
                built.Root = YamlPowerShellValue.ToNode(input, null, out bool truncated);
                if (truncated) WriteWarning("The patch is truncated: objects nested deeper than " + YamlPowerShellValue.MaxDepth + " levels are written as text.");
                report = merge ? document.ApplyMergePatch(built) : document.ApplyJsonPatch(built);
            }
        } catch (PipelineStoppedException) { throw; }
        catch (Exception error) {
            WriteError(OfficeYamlDocuments.Error(error, input));
            return;
        }
        foreach (OfficeIMO.OfficeDiagnostic diagnostic in report.Diagnostics) WriteVerbose(diagnostic.Code + ": " + diagnostic.Message);
        if (PassThru) WriteObject(report);
    }
}
