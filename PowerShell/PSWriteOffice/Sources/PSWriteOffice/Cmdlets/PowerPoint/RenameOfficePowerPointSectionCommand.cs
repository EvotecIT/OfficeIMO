using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Renames a PowerPoint section.</summary>
/// <example>
///   <summary>Rename a section in a presentation.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointRenameSection.pptx -NoSave
/// Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
/// Add-OfficePowerPointSection -Presentation $ppt -Name 'Results' -StartSlideIndex 0
/// Rename-OfficePowerPointSection -Presentation $ppt -Name 'Results' -NewName 'Deep Dive' -PassThru
/// $ppt | Close-OfficePowerPoint -Save</code>
///   <para>Renames the first matching section and returns the updated section metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Rename, "OfficePowerPointSection")]
[OutputType(typeof(PowerPointSectionInfo), typeof(bool))]
public sealed class RenameOfficePowerPointSectionCommand : PSCmdlet {
    /// <summary>Presentation to update (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Existing section name.</summary>
    [Parameter(Mandatory = true)]
    public string Name { get; set; } = string.Empty;

    /// <summary>New section name.</summary>
    [Parameter(Mandatory = true)]
    public string NewName { get; set; } = string.Empty;

    /// <summary>Use case-sensitive matching for the existing section name.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <summary>Emit the renamed section instead of no output.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            bool renamed = presentation.RenameSection(Name, NewName, ignoreCase: !CaseSensitive.IsPresent);
            if (!renamed) {
                WriteError(new ErrorRecord(
                    new InvalidOperationException($"Section '{Name}' was not found."),
                    "PowerPointSectionNotFound",
                    ErrorCategory.ObjectNotFound,
                    Name));
                return;
            }

            if (PassThru.IsPresent) {
                var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                var section = presentation.GetSections().FirstOrDefault(s => string.Equals(s.Name, NewName, comparison));
                if (!string.IsNullOrEmpty(section.Name)) {
                    WriteObject(section);
                } else {
                    WriteObject(true);
                }
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointRenameSectionFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
