using System;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTextRun {
    /// <summary>Gets or sets explicit additional character spacing in points. Values are stored in hundredths of a point; null removes the override.</summary>
    public double? CharacterSpacingPoints {
        get => RunProperties?.Spacing?.Value / 100d;
        set {
            if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value)
                || value.Value is < -4000d or > 4000d))
                throw new ArgumentOutOfRangeException(nameof(value), "Character spacing must be between -4000 and 4000 points.");
            EnsureRunProperties().Spacing = value.HasValue
                ? (int)Math.Round(value.Value * 100d, MidpointRounding.AwayFromZero) : null;
        }
    }
}
