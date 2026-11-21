namespace OfficeIMO;

/// <summary>
/// Controls what an export does with native chart formatting that the shared chart renderer
/// cannot reproduce exactly.
/// </summary>
public enum OfficeChartFidelity {
    /// <summary>
    /// Draws a chart only when its native formatting can be reproduced; any other chart is
    /// skipped and reported with an <c>unsupported-chart</c> warning. This is the default.
    /// </summary>
    Strict,

    /// <summary>
    /// Draws the closest approximation of every chart the renderer can read and reports each
    /// approximation as a <c>chart-approximation</c> warning (<see cref="OfficeDiagnosticCodes.ChartApproximation"/>)
    /// with an approximation loss kind, so callers can still reject the result.
    /// </summary>
    BestEffort
}
