using AngleSharp;
using AngleSharp.Dom;
using Jint;
using Jint.Native;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed partial class RuntimeFrameRealms {
    private readonly HashSet<IWindow> _rootWindows = new(ReferenceEqualityComparer.Instance);
    private IWindow? _rootAnchor;
    private IWindow? _currentRoot;

    internal IWindow ResolveWindow(IWindow window) => _rootWindows.Contains(window) ? _currentRoot ?? window : window;

    internal JsValue WrapWindow(Engine engine, IWindow caller, IWindow? window) {
        if (window == null) return JsValue.Null;
        window = ResolveWindow(window);
        if (ReferenceEquals(window, caller)) return engine.Global;
        return JsValue.FromObject(engine, _rootWindows.Contains(window) ? _rootAnchor! : window);
    }

    internal CancellationToken ResourceLifetimeFor(IBrowsingContext context) {
        lock (sync) {
            for (var current = context; current != null; current = current.Parent) {
                if (_contexts.TryGetValue(current, out var realm)) return realm.Lifetime.Token;
                if (_blockedContexts.Contains(current)) return new CancellationToken(canceled: true);
            }
            return CancellationToken.None;
        }
    }

    internal IReadOnlyList<Engine> RetireRoot() {
        lock (sync) {
            var retained = new HashSet<IBrowsingContext>(ReferenceEqualityComparer.Instance);
            foreach (var window in _auxiliaryWindows.Keys.Where(window => !window.IsClosed)) {
                retained.Add(window.Document.Context);
                CollectAttachedContexts(window.Document, retained);
            }
            var retiring = _realms.Values.Where(realm => !retained.Contains(realm.Context)).ToArray();
            foreach (var realm in retiring) {
                if (IsFrameContext(realm.Context)) realm.Window.Close();
                Retire(realm, blockContext: true, RetirementKind.Replacement, force: true);
            }
            _currentRoot = null;
            return retiring.Select(realm => realm.Engine).ToArray();
        }
    }
}
