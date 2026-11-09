using OfficeIMO.Workflows;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using System.Management.Automation;
using OfficeIMO.OneNote;

namespace PSWriteOffice.Cmdlets.OneNote;

/// <summary>Reads an offline OneNote section, notebook hierarchy, or packaged notebook.</summary>
/// <example>
///   <summary>Read a OneNote section and inspect its pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$section = Open-OfficeOneNote -Path .\Operations.one
/// $section.Pages | Select-Object Title, CreatedUtc, LastModifiedUtc</code>
/// </example>
[Cmdlet(VerbsCommon.Open, "OfficeOneNote", DefaultParameterSetName = "Path")]
[OutputType(typeof(OneNoteSection), typeof(OneNoteNotebook))]
public sealed class OpenOfficeOneNoteCommand : PSCmdlet
{
    /// <summary>Path to a .one section, .onetoc2 notebook index, or .onepkg archive.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = "Path")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Exact input path without wildcard expansion.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "LiteralPath")]
    [Alias("PSPath", "FullName")]
    public string LiteralPath { get; set; } = string.Empty;


    /// <summary>Bounded section and revision-store read options.</summary>
    [Parameter]
    public object[]? Options { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        foreach (var path in OfficePath.ResolveInput(SessionState, ParameterSetName == "LiteralPath" ? LiteralPath : Path, ParameterSetName == "LiteralPath", allowDirectory: true)) ReadPath(path);
    }
    private void ReadPath(string inputPath) {
        var path = inputPath;
        var settings = new OfficeWorkflowRouteSettings();
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        foreach (var input in Options ?? Array.Empty<object>()) {
            var slots = new[] {
                ((object)settings, typeof(OfficeWorkflowRouteSettings).GetProperty(nameof(OfficeWorkflowRouteSettings.OneNoteLoad))!),
                ((object)settings, typeof(OfficeWorkflowRouteSettings).GetProperty(nameof(OfficeWorkflowRouteSettings.OneNoteNotebookLoad))!)
            };
            OfficeOptionPropertyBinding.Apply(engine, OfficeArgumentValue.Unwrap(input), slots);
            settings = (OfficeWorkflowRouteSettings)settings.Snapshot();
        }
        WriteObject(OneNoteCommandUtilities.Read(path, settings.OneNoteLoad, settings.OneNoteNotebookLoad));
    }
}
