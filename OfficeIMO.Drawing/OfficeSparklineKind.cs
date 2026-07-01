namespace OfficeIMO.Drawing;

/// <summary>
/// Shared sparkline visual shapes supported by dependency-free Drawing renderers.
/// </summary>
public enum OfficeSparklineKind {
    /// <summary>Connected line sparkline.</summary>
    Line,

    /// <summary>Column sparkline with value-proportional bars.</summary>
    Column,

    /// <summary>Win/loss sparkline with fixed-height positive and negative bars.</summary>
    WinLoss
}
