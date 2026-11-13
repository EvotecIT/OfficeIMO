using System.Globalization;

namespace OfficeIMO.Toml.Tests.Serialization;

/// <summary>Compares the document model's tree with Tomlyn's model of the same text, value by value.</summary>
internal static class TomlInterop {
    public static bool Same(TomlNode ours, object? theirs, out string difference) {
        difference = string.Empty;
        return Compare(ours, theirs, "$", ref difference);
    }

    private static bool Compare(TomlNode ours, object? theirs, string path, ref string difference) {
        switch (ours) {
            case TomlTable table when theirs is global::Tomlyn.Model.TomlTable other:
                if (table.Count != other.Count) return Fail(path + " has " + other.Count + " keys, not " + table.Count, ref difference);
                foreach (KeyValuePair<string, TomlNode> entry in table) {
                    if (!other.TryGetValue(entry.Key, out object? value)) return Fail(path + " has no key '" + entry.Key + "'", ref difference);
                    if (!Compare(entry.Value, value, path + "." + entry.Key, ref difference)) return false;
                }
                return true;
            case TomlTableArray tables when theirs is global::Tomlyn.Model.TomlTableArray other:
                if (tables.Count != other.Count) return Fail(path + " has " + other.Count + " tables, not " + tables.Count, ref difference);
                for (int i = 0; i < tables.Count; i++) {
                    if (!Compare(tables[i], other[i], path + "[" + i + "]", ref difference)) return false;
                }
                return true;
            case TomlArray array when theirs is global::Tomlyn.Model.TomlArray other:
                if (array.Count != other.Count) return Fail(path + " has " + other.Count + " items, not " + array.Count, ref difference);
                for (int i = 0; i < array.Count; i++) {
                    if (!Compare(array[i], other[i], path + "[" + i + "]", ref difference)) return false;
                }
                return true;
            case TomlValue value:
                return SameValue(value.Value, theirs) || Fail(path + " is " + Describe(theirs) + ", not " + value, ref difference);
            default:
                return Fail(path + " is " + Describe(theirs) + ", not " + ours.Kind, ref difference);
        }
    }

    // Tomlyn parses floats with the runtime, and .NET Framework reads "-0.0" as positive zero. The round trip through
    // OfficeIMO checks the sign on every runtime.
#if NETFRAMEWORK
    private const bool TomlynKeepsZeroSign = false;
#else
    private const bool TomlynKeepsZeroSign = true;
#endif

    private static bool SameValue(object ours, object? theirs) => ours switch {
        double number when theirs is double other => number.Equals(other) && (number != 0 || !TomlynKeepsZeroSign || BitConverter.DoubleToInt64Bits(number) == BitConverter.DoubleToInt64Bits(other)),
        TomlDateTime dateTime when theirs is global::Tomlyn.TomlDateTime other => SameDateTime(dateTime, other),
        _ => Equals(ours, theirs)
    };

    private static bool SameDateTime(TomlDateTime ours, global::Tomlyn.TomlDateTime theirs) {
        // Tomlyn keeps 100-nanosecond ticks; compare at that precision.
        DateTimeOffset value = theirs.DateTime;
        return ours.Kind switch {
            TomlDateTimeKind.OffsetDateTime => ours.ToDateTimeOffset() == value && ours.Offset == value.Offset,
            TomlDateTimeKind.LocalTime => ours.ToTimeOfDay() == value.TimeOfDay,
            _ => ours.ToDateTime() == value.DateTime
        };
    }

    private static string Describe(object? value) => value switch {
        null => "missing",
        IFormattable formattable => value.GetType().Name + " " + formattable.ToString(null, CultureInfo.InvariantCulture),
        _ => value.GetType().Name + " " + value
    };

    private static bool Fail(string message, ref string difference) {
        difference = message;
        return false;
    }
}
