using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a process step to the current designed slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointStep")]
[Alias("PowerPointStep")]
public sealed class AddOfficePowerPointStepCommand : OfficeCmdlet {
    /// <summary>Gets or sets the title.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;
    /// <summary>Gets or sets the text.</summary>
    [Parameter(Position = 1)]
    public string? Text { get; set; }
    /// <summary>Gets or sets the number.</summary>
    [Parameter]
    public string? Number { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try { PowerPointContentScope.Add(new PowerPointLayoutStep(Title, Text, Number)); }
        catch (ArgumentException error) { InvalidArgument(error, this); }
    }
}
