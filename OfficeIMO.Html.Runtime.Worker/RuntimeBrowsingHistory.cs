using Jint.Native;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed record RuntimeNavigation(Uri Url, bool Replace, int EntryIndex = -1, bool Reload = false, int DocumentId = -1);

// Stored graphs contain cloned data only. On document replacement the new realm
// clones every retained state and replaces Entries before executing page script;
// historical state must not keep an old interpreter and DOM alive.
internal sealed class RuntimeBrowsingHistory {
    internal JsValue Entries { get; set; } = JsValue.Null;
    internal int Index { get; set; }
    internal int Generation { get; set; }
    internal RuntimeNavigation? Transition { get; set; }
}
