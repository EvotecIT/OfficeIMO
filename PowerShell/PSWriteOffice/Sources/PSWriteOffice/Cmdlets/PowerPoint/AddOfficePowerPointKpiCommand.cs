using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a metric to the current designed slide.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointKpi", DefaultParameterSetName = "Value")]
[Alias("PowerPointKpi")]
public sealed class AddOfficePowerPointKpiCommand : OfficeCmdlet {
    /// <summary>Gets or sets the value.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Value { get; set; } = string.Empty;
    /// <summary>Gets or sets the label.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Label { get; set; } = string.Empty;
    /// <summary>Gets or sets the change.</summary>
    [Parameter(ParameterSetName = "Value")]
    public string? Change { get; set; }
    /// <summary>Gets or sets an explicit change direction.</summary>
    [Parameter(ParameterSetName = "Value")]
    public PowerPointKpiTrend? Trend { get; set; }
    /// <summary>Gets or sets which direction counts as an improvement.</summary>
    [Parameter]
    public PowerPointKpiDirection GoodDirection { get; set; } = PowerPointKpiDirection.Up;
    /// <summary>Gets or sets the numeric change; percentages are fractional values.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Delta")]
    public double Delta { get; set; }
    /// <summary>Gets or sets the change unit, such as %, pt or ms.</summary>
    [Parameter(ParameterSetName = "Delta")]
    public string? DeltaUnit { get; set; }
    /// <summary>Gets or sets the color.</summary>
    [Parameter, OfficeColorTransformation]
    public OfficeColor? Color { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var metric = ParameterSetName == "Delta"
                ? PowerPointLayoutKpi.FromDelta(Value, Label, Delta, DeltaUnit, GoodDirection, Color)
                : new PowerPointLayoutKpi(Value, Label, Change, Color) { Trend = Trend, GoodDirection = GoodDirection };
            PowerPointContentScope.Add(metric);
        }
        catch (ArgumentException error) { InvalidArgument(error, this); }
    }
}
