using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Gets PowerPoint sections from a presentation.</summary>
/// <para>Returns OfficeIMO section metadata so scripts can inspect section names and slide ranges.</para>
/// <example>
///   <summary>List all sections in a deck.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$ppt = New-OfficePowerPoint -Path .\Examples\Documents\PowerPointSectionsRead.pptx -NoSave
/// Add-OfficePowerPointSlide -Presentation $ppt -Layout 1
/// Add-OfficePowerPointSection -Presentation $ppt -Name 'Appendix' -StartSlideIndex 0
/// Get-OfficePowerPointSection -Presentation $ppt | Select-Object Name, FirstSlideIndex, SlideCount
/// $ppt | Close-OfficePowerPoint</code>
///   <para>Returns section information including section names and slide indexes.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSection")]
[OutputType(typeof(PowerPointSectionInfo))]
public sealed class GetOfficePowerPointSectionCommand : PSCmdlet {
    /// <summary>Presentation to inspect (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointPresentation? Presentation { get; set; }

    /// <summary>Optional section name filter.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Use case-sensitive matching for section names.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var presentation = Presentation ?? PowerPointDslContext.Current?.Presentation
                ?? throw new InvalidOperationException("Presentation was not provided. Use -Presentation or run inside New-OfficePowerPoint.");

            var comparison = CaseSensitive.IsPresent ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (var section in presentation.GetSections()) {
                if (!string.IsNullOrWhiteSpace(Name) &&
                    !string.Equals(section.Name, Name, comparison)) {
                    continue;
                }

                WriteObject(section);
            }
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "PowerPointGetSectionFailed", ErrorCategory.InvalidOperation, Presentation));
        }
    }
}
