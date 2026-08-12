using System.Diagnostics;
using System.Globalization;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Parses and applies machine-specific processor-affinity masks for opt-in benchmark runs.
/// </summary>
internal static class BenchmarkProcessorAffinity {
    /// <summary>Parses a comma-separated list of nonzero decimal or hexadecimal masks.</summary>
    internal static IntPtr[] ParseList(string value) {
        string[] values = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (values.Length == 0) {
            throw new ArgumentException("At least one processor-affinity mask is required.", nameof(value));
        }

        var masks = new IntPtr[values.Length];
        for (int index = 0; index < values.Length; index++) {
            masks[index] = Parse(values[index]);
        }
        return masks;
    }

    /// <summary>Applies one nonzero decimal or hexadecimal mask to the current Windows process.</summary>
    internal static string Apply(string value) {
        if (!OperatingSystem.IsWindows()) {
            throw new PlatformNotSupportedException("Processor-affinity comparison is available only on Windows.");
        }

        IntPtr mask = Parse(value);
        Process.GetCurrentProcess().ProcessorAffinity = mask;
        return Format(mask);
    }

    /// <summary>Applies a non-realtime priority class to the current benchmark process.</summary>
    internal static string ApplyPriority(string value) {
        if (!Enum.TryParse(value, ignoreCase: true, out ProcessPriorityClass priority) ||
            priority == ProcessPriorityClass.RealTime) {
            throw new ArgumentException("Priority must be Idle, BelowNormal, Normal, AboveNormal, or High.", nameof(value));
        }

        using Process process = Process.GetCurrentProcess();
        process.PriorityClass = priority;
        return process.PriorityClass.ToString();
    }

    /// <summary>Formats a mask in the canonical hexadecimal form stored with benchmark evidence.</summary>
    internal static string Format(IntPtr mask) => IntPtr.Size == 8
        ? $"0x{unchecked((ulong)mask.ToInt64()):X}"
        : $"0x{unchecked((uint)mask.ToInt32()):X}";

    private static IntPtr Parse(string value) {
        bool hexadecimal = value.StartsWith("0x", StringComparison.OrdinalIgnoreCase);
        string digits = hexadecimal ? value[2..] : value;
        if (!ulong.TryParse(
                digits,
                hexadecimal ? NumberStyles.AllowHexSpecifier : NumberStyles.None,
                CultureInfo.InvariantCulture,
                out ulong parsed) ||
            parsed == 0 ||
            (IntPtr.Size == 4 && parsed > uint.MaxValue)) {
            throw new ArgumentException(
                $"Processor-affinity mask '{value}' is not a nonzero pointer-sized value.",
                nameof(value));
        }

        return IntPtr.Size == 8
            ? new IntPtr(unchecked((long)parsed))
            : new IntPtr(unchecked((int)parsed));
    }
}
