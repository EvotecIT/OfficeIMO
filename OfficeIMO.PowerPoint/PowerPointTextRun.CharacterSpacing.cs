using System;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTextRun {
    /// <summary>Gets or sets explicit additional physical character spacing. Values are stored in hundredths of a point; null removes the override.</summary>
    public global::OfficeIMO.OfficeLength? CharacterSpacing {
        get => RunProperties?.Spacing?.Value is int spacing
            ? global::OfficeIMO.OfficeLength.Points(spacing / 100D) : (global::OfficeIMO.OfficeLength?)null;
        set {
            double? points = value?.ToPoints();
            if (points.HasValue && (points.Value is < -4000d or > 4000d))
                throw new ArgumentOutOfRangeException(nameof(value), "Character spacing must be between -4000 and 4000 points.");
            EnsureRunProperties().Spacing = value.HasValue
                ? (int)Math.Round(points!.Value * 100d, MidpointRounding.AwayFromZero) : null;
        }
    }
}
