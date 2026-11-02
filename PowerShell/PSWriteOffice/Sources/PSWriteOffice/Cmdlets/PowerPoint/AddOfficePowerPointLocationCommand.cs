using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a named location to a coverage slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointLocation", DefaultParameterSetName = "Location")]
[Alias("PowerPointLocation")]
public sealed class AddOfficePowerPointLocationCommand : OfficeCmdlet {
    /// <summary>Gets or sets the location name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;
    /// <summary>Gets or sets the horizontal fraction of the map.</summary>
    [Parameter(Mandatory = true), ValidateRange(0D, 1D)]
    public double MapLeft { get; set; }
    /// <summary>Gets or sets the vertical fraction of the map.</summary>
    [Parameter(Mandatory = true), ValidateRange(0D, 1D)]
    public double MapTop { get; set; }
    /// <summary>Gets or sets the location detail.</summary>
    [Parameter]
    public string? Detail { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        try { PowerPointContentScope.Add(new PowerPointLayoutLocation(Name, MapLeft, MapTop, Detail)); }
        catch (ArgumentException error) { InvalidArgument(error, Name); }
    }
}
