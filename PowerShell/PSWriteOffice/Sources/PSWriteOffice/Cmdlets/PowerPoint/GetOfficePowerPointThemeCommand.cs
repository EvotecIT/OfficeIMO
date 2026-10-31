using System;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Reads the selected slide master's theme, colors and fonts.</summary>
/// <example><summary>Inspect the first master's theme.</summary>
/// <code>$deck = Open-OfficePowerPoint -Path ./Review.pptx
/// Get-OfficePowerPointTheme -Document $deck -Master 1
/// Close-OfficePowerPoint -Document $deck</code></example>
/// <example><summary>Inspect the theme while composing a deck.</summary>
/// <code>New-OfficePowerPoint -Path ./Review.pptx -Theme Modern {
///     $theme = Get-OfficePowerPointTheme
///     PowerPointSlide -Title 'Service review'
/// }</code></example>
[Cmdlet(VerbsCommon.Get, "OfficePowerPointTheme")]
[OutputType(typeof(PowerPointThemeInfo))]
public sealed class GetOfficePowerPointThemeCommand : OfficeCmdlet {
    /// <summary>Gets or sets the live document; omission uses its enclosing content block.</summary>
    [Parameter(Position = 0, ValueFromPipeline = true)]
    public PowerPointDocument? Document { get; set; }
    /// <summary>Gets or sets a one-based master position; defaults to the first.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int Master { get; set; } = 1;
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var document = ResolveDocument(Document);
        try {
            int index = Master - 1;
            var colors = document.GetThemeColors(index);
            var fonts = document.GetThemeFonts(index);
            string name = PowerPointTemplate.Inspect(document).Masters.Single(master => master.MasterIndex == index).ThemeName;
            WriteObject(new PowerPointThemeInfo(index, name, colors, fonts));
        } catch (ArgumentException error) { InvalidArgument(error, document); }
    }
}
