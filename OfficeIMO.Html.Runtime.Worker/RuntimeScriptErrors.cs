using Jint;
using Jint.Native;
using Jint.Native.Promise;

namespace OfficeIMO.Html.Runtime.Worker;

// Rejections are checked at script-turn boundaries, allowing handlers added in the same turn.
internal sealed class RuntimeScriptErrors {
    private readonly Dictionary<object, JsValue> _rejections = new(ReferenceEqualityComparer.Instance);
    private readonly int _maximum;
    private string? _firstError;

    internal RuntimeScriptErrors(int maximum) => _maximum = maximum;

    internal void Attach(Engine engine) => engine.Advanced.PromiseRejectionTracker += (_, args) => {
        if (args.Operation == PromiseRejectionOperation.Handle) _rejections.Remove(args.Promise);
        else if (_rejections.Count >= _maximum) Report("Pending promise rejections exceeded their tracking budget.");
        else _rejections[args.Promise] = args.Value ?? JsValue.Undefined;
    };

    internal void Report(string message) => Interlocked.CompareExchange(ref _firstError, message, null);

    internal void ThrowIfFailed(bool includeRejections = true) {
        if (_firstError is string error) throw new HtmlScriptRuntimeException(error);
        if (includeRejections && _rejections.Count != 0) throw new HtmlScriptRuntimeException("Unhandled promise rejection: " + _rejections.Values.First());
    }
}
