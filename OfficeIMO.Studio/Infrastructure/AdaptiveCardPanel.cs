using Avalonia;
using Avalonia.Controls.Primitives;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Arranges task cards into equal columns that retain a readable minimum width.</summary>
public sealed class AdaptiveCardPanel : UniformGrid {
    /// <summary>Defines the preferred minimum width of a card before another row is used.</summary>
    public static readonly StyledProperty<double> MinimumCardWidthProperty =
        AvaloniaProperty.Register<AdaptiveCardPanel, double>(nameof(MinimumCardWidth), 260D,
            validate: value => double.IsFinite(value) && value > 0D);

    /// <summary>Gets or sets the preferred minimum card width in device-independent pixels.</summary>
    public double MinimumCardWidth {
        get => GetValue(MinimumCardWidthProperty);
        set => SetValue(MinimumCardWidthProperty, value);
    }

    static AdaptiveCardPanel() => AffectsMeasure<AdaptiveCardPanel>(MinimumCardWidthProperty);

    protected override Size MeasureOverride(Size availableSize) {
        int count = Math.Max(1, Children.Count(child => child.IsVisible));
        int columns = double.IsFinite(availableSize.Width)
            ? Math.Clamp((int)Math.Floor((availableSize.Width + ColumnSpacing) /
                (MinimumCardWidth + ColumnSpacing)), 1, count)
            : count;
        // Balance the last row instead of leaving one of six tasks on its own.
        int rows = (int)Math.Ceiling((double)count / columns);
        columns = (int)Math.Ceiling((double)count / rows);
        if (Columns != columns) SetCurrentValue(ColumnsProperty, columns);
        return base.MeasureOverride(availableSize);
    }
}
