using System.Management.Automation;
using OfficeIMO.OneNote;

namespace PSWriteOffice.Cmdlets.OneNote;

/// <summary>Reads an offline OneNote section, notebook hierarchy, or packaged notebook.</summary>
/// <example>
///   <summary>Read a OneNote section and inspect its pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$section = Get-OfficeOneNote -Path .\Operations.one
/// $section.Pages | Select-Object Title, CreatedUtc, LastModifiedUtc</code>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeOneNote")]
[OutputType(typeof(OneNoteSection), typeof(OneNoteNotebook))]
public sealed class GetOfficeOneNoteCommand : PSCmdlet
{
    /// <summary>Path to a .one section, .onetoc2 notebook index, or .onepkg archive.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Bounded section and revision-store read options.</summary>
    [Parameter]
    public OneNoteReaderOptions? Options { get; set; }

    /// <summary>Notebook hierarchy, package, and section-error policy.</summary>
    [Parameter]
    public OneNoteNotebookReaderOptions? NotebookOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        WriteObject(OneNoteCommandUtilities.Read(path, Options, NotebookOptions));
    }
}
