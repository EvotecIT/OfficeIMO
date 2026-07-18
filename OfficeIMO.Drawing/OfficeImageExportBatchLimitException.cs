using System;

namespace OfficeIMO.Drawing;

/// <summary>Raised when aggregate batch output exceeds a configured safety budget.</summary>
public sealed class OfficeImageExportBatchLimitException : InvalidOperationException {
    internal OfficeImageExportBatchLimitException(string limitName, long actual, long maximum)
        : base("Image export batch exceeded " + limitName + ": " + actual + " requested, " + maximum + " allowed.") {
        LimitName = limitName;
        Actual = actual;
        Maximum = maximum;
    }

    /// <summary>Name of the exceeded budget.</summary>
    public string LimitName { get; }

    /// <summary>Observed aggregate value.</summary>
    public long Actual { get; }

    /// <summary>Configured maximum.</summary>
    public long Maximum { get; }
}
