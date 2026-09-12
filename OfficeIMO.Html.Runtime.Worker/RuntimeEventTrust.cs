using System.Runtime.CompilerServices;
using AngleSharp.Dom.Events;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeEventTrust {
    private sealed class State { internal bool Trusted; }
    private static readonly ConditionalWeakTable<Event, State> States = new();
    internal static void Set(Event value, bool trusted) => States.GetOrCreateValue(value).Trusted=trusted;
    internal static bool Read(Event value) => States.TryGetValue(value,out var state) ? state.Trusted : value.IsTrusted;
}
