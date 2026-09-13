using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeDomSynchronization(Func<object?> engine) : IDomSynchronization {
    public object SyncRoot => engine() ?? this;
}
