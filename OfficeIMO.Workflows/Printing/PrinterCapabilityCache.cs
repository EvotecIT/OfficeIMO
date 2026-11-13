using System.Diagnostics;

namespace OfficeIMO.Workflows;

/// <summary>
/// Printer capabilities a <see cref="PdfPrinterService"/> read recently, so a submission that follows a capability query, or another
/// submission to the same printer, does not ask the printer again. Entries expire quickly and are removed when a submission fails.
/// </summary>
internal sealed class PrinterCapabilityCache {
    internal static readonly TimeSpan DefaultLifetime = TimeSpan.FromMinutes(1);

    private readonly Dictionary<string, (object Value, long Expires)> _entries = new(StringComparer.Ordinal);
    private readonly long _lifetimeTicks;
    private readonly Func<long> _clock;

    public PrinterCapabilityCache() : this(DefaultLifetime, Stopwatch.GetTimestamp) { }

    /// <param name="lifetime">How long an entry is served.</param>
    /// <param name="clock">Current time in <see cref="Stopwatch"/> ticks.</param>
    internal PrinterCapabilityCache(TimeSpan lifetime, Func<long> clock) {
        _lifetimeTicks = (long)(lifetime.TotalSeconds * Stopwatch.Frequency);
        _clock = clock;
    }

    /// <summary>Returns the capabilities stored for the printer when they are of the requested kind and have not expired.</summary>
    public bool TryGet<T>(string printerName, out T value) where T : class {
        lock (_entries) {
            if (_entries.TryGetValue(printerName, out var entry)) {
                if (_clock() >= entry.Expires) _entries.Remove(printerName);
                else if (entry.Value is T stored) {
                    value = stored;
                    return true;
                }
            }
        }
        value = null!;
        return false;
    }

    public void Set(string printerName, object value) {
        lock (_entries) _entries[printerName] = (value, _clock() + _lifetimeTicks);
    }

    public void Remove(string printerName) {
        lock (_entries) _entries.Remove(printerName);
    }
}
