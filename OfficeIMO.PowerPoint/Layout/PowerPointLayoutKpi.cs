using System;
using System.Globalization;

namespace OfficeIMO.PowerPoint {
    /// <summary>The direction of a metric's change indicator.</summary>
    internal enum PowerPointKpiTrend {
        /// <summary>The metric increased.</summary>
        Up,
        /// <summary>The metric decreased.</summary>
        Down,
        /// <summary>The metric stayed unchanged.</summary>
        Flat
    }

    /// <summary>The direction in which a metric improves.</summary>
    internal enum PowerPointKpiDirection {
        /// <summary>Higher values are better.</summary>
        Up,
        /// <summary>Lower values are better.</summary>
        Down
    }

    /// <summary>A formatted metric, label and optional change.</summary>
    internal sealed class PowerPointLayoutKpi : PowerPointLayoutBlock {
        /// <summary>Creates content for the shared layout engine.</summary>
        public PowerPointLayoutKpi(string value, string label, string? change = null, OfficeColor? color = null) {
            Value = value ?? throw new ArgumentNullException(nameof(value));
            Label = label ?? throw new ArgumentNullException(nameof(label));
            Change = change;
            Color = color;
        }

        /// <summary>Formatted metric value.</summary>
        public string Value { get; }
        /// <summary>Metric label.</summary>
        public string Label { get; }
        /// <summary>Optional formatted change.</summary>
        public string? Change { get; }
        /// <summary>Explicit indicator colour; otherwise the trend and good direction choose a semantic colour.</summary>
        public OfficeColor? Color { get; }
        /// <summary>Explicit change direction; omission infers it from a leading plus or minus in Change.</summary>
        public PowerPointKpiTrend? Trend { get; set; }
        /// <summary>The direction counted as an improvement; higher is better by default.</summary>
        public PowerPointKpiDirection GoodDirection { get; set; } = PowerPointKpiDirection.Up;

        /// <summary>Creates a metric with an invariant numeric change rounded to two decimal places and its direction.</summary>
        /// <param name="value">Formatted metric value.</param>
        /// <param name="label">Metric label.</param>
        /// <param name="delta">Numeric change; with the percent unit, a fraction such as 0.12 means 12%.</param>
        /// <param name="unit">Optional unit, such as %, pt or ms.</param>
        /// <param name="goodDirection">Direction counted as an improvement.</param>
        /// <param name="color">Explicit indicator colour overriding semantic colour selection.</param>
        public static PowerPointLayoutKpi FromDelta(string value, string label, double delta, string? unit = null,
            PowerPointKpiDirection goodDirection = PowerPointKpiDirection.Up, OfficeColor? color = null) {
            string suffix = unit?.Trim() ?? string.Empty;
            double amount = suffix == "%" ? delta * 100D : delta;
            if (double.IsNaN(amount) || double.IsInfinity(amount)) throw new ArgumentOutOfRangeException(nameof(delta), "The change must be finite.");
            string change = amount.ToString("+0.##;-0.##;0", CultureInfo.InvariantCulture) +
                (suffix.Length == 0 || suffix == "%" ? suffix : " " + suffix);
            return new PowerPointLayoutKpi(value, label, change, color) {
                Trend = delta > 0D ? PowerPointKpiTrend.Up : delta < 0D ? PowerPointKpiTrend.Down : PowerPointKpiTrend.Flat,
                GoodDirection = goodDirection
            };
        }

        internal PowerPointKpiTrend? EffectiveTrend {
            get {
                if (Trend.HasValue) return Trend;
                string change = Change?.TrimStart() ?? string.Empty;
                if (change.StartsWith("+", StringComparison.Ordinal)) return PowerPointKpiTrend.Up;
                if (change.StartsWith("-", StringComparison.Ordinal) || change.StartsWith("−", StringComparison.Ordinal)) return PowerPointKpiTrend.Down;
                return null;
            }
        }

        internal string ChangeText => (EffectiveTrend switch {
            PowerPointKpiTrend.Up => "▲ ", PowerPointKpiTrend.Down => "▼ ", _ => string.Empty
        }) + (Change?.TrimStart() ?? string.Empty);

        internal override bool IsContentBlock => false;
    }
}
