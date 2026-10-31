using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a narrative card or logo tile to the current designed slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointCard")]
[Alias("PowerPointCard")]
public sealed class AddOfficePowerPointCardCommand : OfficeCmdlet {
    /// <summary>Gets or sets the title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;
    /// <summary>Gets or sets the text.</summary>
    [Parameter]
    public string? Text { get; set; }
    /// <summary>Gets or sets the items.</summary>
    [Parameter]
    public string[]? Items { get; set; }
    /// <summary>Gets or sets the image.</summary>
    [Parameter]
    public string? Image { get; set; }
    /// <summary>Gets or sets the color.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Color { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try { PowerPointContentScope.Add(new PowerPointLayoutCard(Title, Text, Items, Image == null ? null : SessionState.Path.GetUnresolvedProviderPathFromPSPath(Image), Color)); }
        catch (ArgumentException error) { InvalidArgument(error, this); }
    }
}
