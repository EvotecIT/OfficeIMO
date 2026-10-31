using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads sections from a live presentation, optionally matching a PowerShell wildcard pattern.</summary>
/// <example><summary>List the sections in a deck.</summary>
/// <code>$deck = Open-OfficePowerPoint -Path ./Review.pptx
/// Get-OfficePowerPointSection -Document $deck
/// Close-OfficePowerPoint -Document $deck</code></example>
/// <example><summary>Find the appendix section.</summary>
/// <code>$deck | Get-OfficePowerPointSection -Name Appendix</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointSection")]
[OutputType(typeof(PowerPointSectionInfo))]
public sealed class GetOfficePowerPointSectionCommand : OfficeCmdlet {
    /// <summary>Gets or sets the live document; omission uses its enclosing content block.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true)]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets a section name or PowerShell wildcard pattern.</summary>
    [Parameter, SupportsWildcards]
    public string? Name { get; set; }
    /// <summary>Gets or sets whether name matching is case-sensitive.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ResolveDocument(Document);
        var options = CaseSensitive ? WildcardOptions.None : WildcardOptions.IgnoreCase;
        var pattern = string.IsNullOrWhiteSpace(Name) ? null : new WildcardPattern(Name!, options);
        foreach (var section in document.GetSections())
            if (pattern == null || pattern.IsMatch(section.Name)) WriteObject(section);
    }
}
