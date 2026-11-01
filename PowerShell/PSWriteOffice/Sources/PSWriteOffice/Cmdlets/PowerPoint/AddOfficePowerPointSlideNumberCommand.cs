using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Inserts a native slide-number field into the current footer.</summary>
/// <example><summary>Show the current slide and total.</summary>
/// <code>PowerPointFooter -Alignment Right { PowerPointSlideNumber -IncludeTotal }</code></example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointSlideNumber")]
[Alias("PowerPointSlideNumber")]
public sealed class AddOfficePowerPointSlideNumberCommand : OfficeCmdlet {
    /// <summary>Append the physical slide count, refreshed when the deck is saved.</summary>
    [Parameter]
    public SwitchParameter IncludeTotal { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var scope = PowerPointTextScope.Current;
        if (PowerPointFooterScope.Current == null || scope == null)
            throw new PSInvalidOperationException("Invoke PowerPointSlideNumber inside PowerPointFooter.");
        PowerPointDocument.AddFooterNumber(scope.TextParagraph(), IncludeTotal);
    }
}
