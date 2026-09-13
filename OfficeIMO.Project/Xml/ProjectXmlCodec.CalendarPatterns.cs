using System.Globalization;
using System.Text;

namespace OfficeIMO.Project;

internal static partial class ProjectXmlCodec {
    private static string CalendarPattern(DateTime? from, DateTime? to, bool? working, IEnumerable<ProjectWorkingInterval> intervals, CancellationToken token) {
        var key = new StringBuilder();
        key.Append(from?.Ticks.ToString(CultureInfo.InvariantCulture)).Append('|');
        key.Append(to?.Ticks.ToString(CultureInfo.InvariantCulture)).Append('|');
        key.Append(working.HasValue ? (working.Value ? "1" : "0") : "").Append('|');
        foreach (var interval in intervals) {
            token.ThrowIfCancellationRequested();
            key.Append(interval.From?.Ticks.ToString(CultureInfo.InvariantCulture)).Append(':');
            key.Append(interval.To?.Ticks.ToString(CultureInfo.InvariantCulture)).Append(';');
        }
        return key.ToString();
    }
}
