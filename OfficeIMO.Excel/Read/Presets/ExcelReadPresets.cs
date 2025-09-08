using System;
using System.Globalization;

namespace OfficeIMO.Excel.Read
{
    /// <summary>
    /// Convenience presets for common read conversions to keep user code minimal.
    /// Provides several choices from minimal to more aggressive conversions.
    /// </summary>
    public static class ExcelReadPresets
    {
        /// <summary>
        /// No extra converters. Equivalent to default ExcelReadOptions with only built-in behavior.
        /// </summary>
        public static ExcelReadOptions None()
        {
            return new ExcelReadOptions();
        }

        /// <summary>
        /// Minimal, non-intrusive converters:
        /// - Y/Yes/N/No to bool
        /// - Currency-like strings to decimal (tries CurrentCulture, InvariantCulture, a few common locales)
        /// </summary>
        public static ExcelReadOptions Simple()
        {
            var opt = new ExcelReadOptions();
            ApplyBasicConverters(opt);
            return opt;
        }

        /// <summary>
        /// Like Simple() but prefers decimal for numeric cells when possible.
        /// </summary>
        public static ExcelReadOptions DecimalFirst()
        {
            var opt = new ExcelReadOptions { NumericAsDecimal = true };
            ApplyBasicConverters(opt);
            return opt;
        }

        /// <summary>
        /// Applies the same converters as returned by <see cref="Simple"/> to an existing options instance.
        /// </summary>
        public static void ApplyBasicConverters(ExcelReadOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            options.CellValueConverter = ctx =>
            {
                var s = ctx.RawText?.Trim();
                if (!string.IsNullOrEmpty(s))
                {
                    if (string.Equals(s, "Y", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "Yes", StringComparison.OrdinalIgnoreCase))
                        return new ExcelCellValue(true);
                    if (string.Equals(s, "N", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "No", StringComparison.OrdinalIgnoreCase))
                        return new ExcelCellValue(false);
                }
                return ExcelCellValue.NotHandled;
            };

            options.TypeConverter = (raw, target, culture) =>
            {
                if (target == typeof(decimal) && raw != null)
                {
                    var s = Convert.ToString(raw, culture);
                    if (!string.IsNullOrEmpty(s))
                    {
                        decimal dec;
                        // Try a few likely currency formats/locales
                        var cultures = new[]
                        {
                            CultureInfo.CurrentCulture,
                            CultureInfo.InvariantCulture,
                            CultureInfo.GetCultureInfo("en-US"),
                            CultureInfo.GetCultureInfo("pl-PL"),
                            CultureInfo.GetCultureInfo("de-DE"),
                        };

                        foreach (var ci in cultures)
                        {
                            if (decimal.TryParse(s, NumberStyles.Currency, ci, out dec))
                                return (true, dec);
                        }

                        // Fall back: strip symbols and re-parse
                        var cleaned = s.Replace("$", string.Empty).Replace("€", string.Empty).Replace("PLN", string.Empty).Replace(" ", string.Empty);
                        foreach (var ci in cultures)
                        {
                            if (decimal.TryParse(cleaned, NumberStyles.Number | NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint, ci, out dec))
                                return (true, dec);
                        }
                    }
                }
                return (false, null);
            };
        }

        /// <summary>
        /// Aggressive converters for convenience:
        /// - Extended bool: true/false/1/0/on/off/y/n/yes/no
        /// - Percent strings (e.g., 12%) to decimals (0.12)
        /// - Currency and plain numeric strings to decimal/double/int where appropriate
        /// - ISO date strings to DateTime
        /// </summary>
        public static ExcelReadOptions Aggressive()
        {
            var opt = new ExcelReadOptions();

            opt.CellValueConverter = ctx =>
            {
                var s = ctx.RawText?.Trim();
                if (!string.IsNullOrEmpty(s))
                {
                    // Bool aliases
                    if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) || s == "1" || string.Equals(s, "on", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "y", StringComparison.OrdinalIgnoreCase))
                        return new ExcelCellValue(true);
                    if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase) || s == "0" || string.Equals(s, "off", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "no", StringComparison.OrdinalIgnoreCase) || string.Equals(s, "n", StringComparison.OrdinalIgnoreCase))
                        return new ExcelCellValue(false);

                    // Percent → decimal
                    if (s.EndsWith("%", StringComparison.Ordinal))
                    {
                        var val = s.Substring(0, s.Length - 1);
                        if (double.TryParse(val, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var pInv))
                            return new ExcelCellValue((decimal)(pInv / 100.0));
                        if (double.TryParse(val, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var pCur))
                            return new ExcelCellValue((decimal)(pCur / 100.0));
                    }

                    // ISO date
                    if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var iso))
                        return new ExcelCellValue(iso);
                }
                return ExcelCellValue.NotHandled;
            };

            opt.TypeConverter = (raw, target, culture) =>
            {
                var s = raw is string str ? str.Trim() : null;
                // decimals and doubles
                if ((target == typeof(decimal) || target == typeof(double)) && s != null)
                {
                    foreach (var ci in new[] { CultureInfo.CurrentCulture, CultureInfo.InvariantCulture, CultureInfo.GetCultureInfo("en-US"), CultureInfo.GetCultureInfo("pl-PL"), CultureInfo.GetCultureInfo("de-DE") })
                    {
                        if (target == typeof(decimal) && decimal.TryParse(s, NumberStyles.Any, ci, out var dec))
                            return (true, dec);
                        if (target == typeof(double) && double.TryParse(s, NumberStyles.Any, ci, out var dbl))
                            return (true, dbl);
                    }
                }
                // ints
                if ((target == typeof(int) || target == typeof(long)) && s != null)
                {
                    foreach (var ci in new[] { CultureInfo.CurrentCulture, CultureInfo.InvariantCulture })
                    {
                        if (target == typeof(int) && int.TryParse(s, NumberStyles.Integer, ci, out var i32))
                            return (true, i32);
                        if (target == typeof(long) && long.TryParse(s, NumberStyles.Integer, ci, out var i64))
                            return (true, i64);
                    }
                }
                // ISO date
                if (target == typeof(DateTime) && s != null)
                {
                    if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var dt))
                        return (true, dt);
                }
                return (false, null);
            };

            return opt;
        }
    }
}
