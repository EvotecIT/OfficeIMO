namespace OfficeIMO;

/// <summary>
/// How native Word, Excel and PowerPoint charts are drawn by image and PDF exports.
/// </summary>
/// <remarks>
/// Every export reports a chart it skips with an <c>unsupported-chart</c> warning
/// (<see cref="OfficeDiagnosticCodes.UnsupportedChart"/>; Word PDF and page images use
/// <c>native-body-chart-unsupported</c>) and each approximation with a <c>chart-approximation</c>
/// warning (<see cref="OfficeDiagnosticCodes.ChartApproximation"/>). Formatting that every fidelity draws
/// as an approximation, such as a number format drawn as General, is also reported as <c>chart-approximation</c>.
/// </remarks>
public sealed class OfficeChartRenderOptions {
    /// <summary>
    /// What to do with formatting the shared chart renderer cannot reproduce exactly.
    /// The default is <see cref="OfficeChartFidelity.Strict"/>.
    /// </summary>
    public OfficeChartFidelity Fidelity { get; set; } = OfficeChartFidelity.Strict;

    /// <summary>Creates an independent copy of these settings.</summary>
    public OfficeChartRenderOptions Clone() => new() { Fidelity = Fidelity };

    internal void Validate() {
        if (!Enum.IsDefined(typeof(OfficeChartFidelity), Fidelity))
            throw new ArgumentOutOfRangeException(nameof(Fidelity));
    }
}
