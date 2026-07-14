namespace OfficeIMO.Pdf;

/// <summary>Configures conditional, bounded, and replayable document flow.</summary>
public sealed class PdfFlowOptions {
    private double _minimumRemainingHeight;

    /// <summary>Optional predicate evaluated when the group reaches layout.</summary>
    public Func<PdfFlowContext, bool>? ShowIf { get; set; }

    /// <summary>Minimum remaining height required before the group starts; otherwise a new page is started.</summary>
    public double MinimumRemainingHeight {
        get => _minimumRemainingHeight;
        set {
            Guard.NonNegative(value, nameof(value));
            _minimumRemainingHeight = value;
        }
    }

    /// <summary>Attempts to keep all measurable nested content on one page.</summary>
    public bool KeepTogether { get; set; }

    /// <summary>Behavior used when the measurable group does not fit in the current page remainder.</summary>
    public PdfFlowOverflowBehavior OverflowBehavior { get; set; } = PdfFlowOverflowBehavior.Continue;

    internal PdfFlowOptions Clone() {
        return new PdfFlowOptions {
            ShowIf = ShowIf,
            MinimumRemainingHeight = MinimumRemainingHeight,
            KeepTogether = KeepTogether,
            OverflowBehavior = OverflowBehavior
        };
    }
}
