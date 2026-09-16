using System.Globalization;

namespace OfficeIMO.GoogleWorkspace {
    /// <summary>Formats values used in persisted Google Workspace synchronization fingerprints.</summary>
    /// <remarks>Uses explicit, invariant round-trip precision so a checkpoint created on .NET Framework can be compared on modern .NET runtimes.</remarks>
    public static class GoogleWorkspaceCheckpointFormat {
        /// <summary>Formats an interpolated fingerprint without depending on the current culture or runtime's default floating-point precision.</summary>
        public static string Format(FormattableString value) {
            if (value == null) throw new ArgumentNullException(nameof(value));
            return value.ToString(StableFormatProvider.Instance);
        }

        private sealed class StableFormatProvider : IFormatProvider, ICustomFormatter {
            internal static readonly StableFormatProvider Instance = new StableFormatProvider();

            public object? GetFormat(Type? formatType) => formatType == typeof(ICustomFormatter)
                ? this
                : CultureInfo.InvariantCulture.GetFormat(formatType);

            public string Format(string? format, object? value, IFormatProvider? provider) {
                if (value == null) return string.Empty;
                if (value is double number) return number.ToString(format ?? "G17", CultureInfo.InvariantCulture);
                if (value is float single) return single.ToString(format ?? "G9", CultureInfo.InvariantCulture);
                if (value is decimal decimalNumber) return decimalNumber.ToString(format ?? "G29", CultureInfo.InvariantCulture);
                if (value is DateTime date) return date.ToString(format ?? "O", CultureInfo.InvariantCulture);
                if (value is DateTimeOffset offset) return offset.ToString(format ?? "O", CultureInfo.InvariantCulture);
                if (value is TimeSpan duration) return duration.ToString(format ?? "c", CultureInfo.InvariantCulture);
                return value is IFormattable formattable
                    ? formattable.ToString(format, CultureInfo.InvariantCulture) ?? string.Empty
                    : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }
    }
}
