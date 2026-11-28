namespace OfficeIMO;

/// <summary>A native owner's inert source callback. It is internal so consumers cannot replace native document discovery.</summary>
internal sealed class OfficeEmbeddedPayloadSource {
    internal OfficeEmbeddedPayloadSource(OfficeEmbeddedPayloadInfo info, Func<CancellationToken, Stream> openRead,
        Action<Action>? executeRead = null, Action<Action, CancellationToken>? executeCancelableRead = null) {
        Info = info ?? throw new ArgumentNullException(nameof(info));
        OpenRead = openRead ?? throw new ArgumentNullException(nameof(openRead));
        ExecuteRead = executeCancelableRead ?? (executeRead == null ? static (action, _) => action() : (action, _) => executeRead(action));
    }
    internal OfficeEmbeddedPayloadInfo Info { get; }
    internal Func<CancellationToken, Stream> OpenRead { get; }
    internal Action<Action, CancellationToken> ExecuteRead { get; }
}
