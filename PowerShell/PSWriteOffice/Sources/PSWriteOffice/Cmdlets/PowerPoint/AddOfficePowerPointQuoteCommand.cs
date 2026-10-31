using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a quotation and attribution to the current designed slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointQuote")]
[Alias("PowerPointQuote")]
public sealed class AddOfficePowerPointQuoteCommand : OfficeCmdlet {
    /// <summary>Gets or sets the text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;
    /// <summary>Gets or sets the attribution.</summary>
    [Parameter(Position = 1)]
    public string? Attribution { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try { PowerPointContentScope.Add(new PowerPointLayoutQuote(Text, Attribution)); }
        catch (ArgumentException error) { InvalidArgument(error, this); }
    }
}
