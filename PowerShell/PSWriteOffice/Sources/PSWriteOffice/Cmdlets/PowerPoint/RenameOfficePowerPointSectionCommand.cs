using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Renames the first section matching the supplied name.</summary>
/// <example><summary>Rename a section in a live deck.</summary>
/// <code>Rename-OfficePowerPointSection -Document $deck -Name Results -NewName 'Deep dive'</code></example>
/// <example><summary>Return the renamed section.</summary>
/// <code>$section = Rename-OfficePowerPointSection -Document $deck -Name Appendix -NewName Reference -PassThru</code></example>
[Cmdlet(VerbsCommon.Rename, "OfficePowerPointSection", DefaultParameterSetName = "Document")]
[OutputType(typeof(PowerPointSectionInfo))]
public sealed class RenameOfficePowerPointSectionCommand : OfficeCmdlet {
    /// <summary>Gets or sets the live document; omission uses its enclosing content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets the existing section name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;
    /// <summary>Gets or sets the new section name.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string NewName { get; set; } = string.Empty;
    /// <summary>Gets or sets whether name matching is case-sensitive.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }
    /// <summary>Gets or sets whether the changed section is returned.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ResolveDocument(Document);
        try {
            var comparison = CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            var selected = document.GetSections().FirstOrDefault(section => string.Equals(section.Name, Name, comparison));
            if (string.IsNullOrEmpty(selected.Id) || !document.RenameSection(Name, NewName, ignoreCase: !CaseSensitive)) {
                WriteError(OfficeError.Create(new ItemNotFoundException("Section '" + Name + "' was not found."),
                    "section-not-found", Name, ErrorCategory.ObjectNotFound));
                return;
            }
            if (PassThru) WriteObject(document.GetSections().Single(section => section.Id == selected.Id));
        } catch (ArgumentException error) { InvalidArgument(error, document); }
    }
}
