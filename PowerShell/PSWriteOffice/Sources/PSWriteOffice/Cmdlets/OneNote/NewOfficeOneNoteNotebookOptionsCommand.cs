using System.Management.Automation;
using OfficeIMO.OneNote;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.OneNote;

/// <summary>Creates native OneNoteNotebookReaderOptions settings for document operations.</summary>
/// <example><summary>Create default settings.</summary><code>$options = New-OfficeOneNoteNotebookOptions</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOneNoteNotebookOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(OfficeIMO.OneNote.OneNoteNotebookReaderOptions))]
public sealed class NewOfficeOneNoteNotebookOptionsCommand : OfficeCmdlet {
    /// <summary>NoSectionContent override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoSectionContent { get; set; }
    /// <summary>StopOnSectionError override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter StopOnSectionError { get; set; }
    /// <summary>NoSectionGroupRecursion override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter NoSectionGroupRecursion { get; set; }
    /// <summary>IncludeRecycleBin override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public SwitchParameter IncludeRecycleBin { get; set; }
    /// <summary>MaxSectionGroupDepth override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaxSectionGroupDepth { get; set; }
    /// <summary>MaxNotebookEntries override; omission preserves the native default.</summary>
    [Parameter(ParameterSetName = "Default")] public int? MaxNotebookEntries { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new OfficeIMO.OneNote.OneNoteNotebookReaderOptions();
        if (BoundSwitch(nameof(NoSectionContent)) is bool valueNoSectionContent) options.LoadSectionContent = !valueNoSectionContent;
        if (BoundSwitch(nameof(StopOnSectionError)) is bool valueStopOnSectionError) options.ContinueOnSectionError = !valueStopOnSectionError;
        if (BoundSwitch(nameof(NoSectionGroupRecursion)) is bool valueNoSectionGroupRecursion) options.RecurseSectionGroups = !valueNoSectionGroupRecursion;
        if (BoundSwitch(nameof(IncludeRecycleBin)) is bool valueIncludeRecycleBin) options.IncludeRecycleBin = valueIncludeRecycleBin;
        if (MaxSectionGroupDepth != null) options.MaxSectionGroupDepth = MaxSectionGroupDepth.Value;
        if (MaxNotebookEntries != null) options.MaxNotebookEntries = MaxNotebookEntries.Value;
        WriteObject(options);
    }
}
