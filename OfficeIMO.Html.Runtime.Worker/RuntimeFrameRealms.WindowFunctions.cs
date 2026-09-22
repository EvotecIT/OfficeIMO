using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Interop;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Html.Runtime.Worker;

// The provider's Window proxy exposes native DOM members, while script-defined
// properties live on a different Jint global. Bridge only functions added after
// realm initialization; each call rechecks origin and the exact realm lifetime.
internal sealed partial class RuntimeFrameRealms {
    private readonly ConditionalWeakTable<Engine, ConditionalWeakTable<IWindow, ObjectInstance>> _windowFunctionViews = new();

    internal void CaptureWindowSurface(IDocument document) {
        lock (sync) {
            if (document.DefaultView is { } window && _realms.TryGetValue(window, out var realm))
                realm.InitialGlobalKeys = realm.Engine.Global.GetOwnProperties()
                    .Where(property => property.Key.IsString())
                    .Select(property => property.Key.AsString())
                    .ToHashSet(StringComparer.Ordinal);
        }
    }

    private JsValue WrapWindowFunctions(Engine engine, IWindow caller, IWindow target) =>
        _windowFunctionViews.GetValue(engine, _ => new ConditionalWeakTable<IWindow, ObjectInstance>())
            .GetValue(target, value => CreateWindowFunctionView(engine, caller, value));

    private ObjectInstance CreateWindowFunctionView(Engine engine, IWindow caller, IWindow target) {
        var native = JsValue.FromObject(engine, target).AsObject();
        var cache = new Dictionary<string, (Realm Realm, Function Function, JsValue Forwarder)>(StringComparer.Ordinal);
        var lookup = new ClrFunction(engine, "getWindowFunction", (_, args) => {
            if (!TryWindowFunction(engine, caller, target, args[0], out var realm, out var function)) return JsValue.Undefined;
            string name = args[0].AsString();
            if (cache.TryGetValue(name, out var cached) && ReferenceEquals(cached.Realm, realm) &&
                ReferenceEquals(cached.Function, function)) return cached.Forwarder;
            var forwarder = new ClrFunction(engine, name, (receiver, callArgs) =>
                InvokeWindowFunction(engine, caller, target, realm, function, receiver, callArgs));
            cache[name] = (realm, function, forwarder);
            return forwarder;
        });
        var has = new ClrFunction(engine, "hasWindowFunction", (receiver, args) =>
            TryWindowFunction(engine, caller, target, args[0], out _, out _));
        var factory = engine.Evaluate("(target,lookup,has)=>new Proxy(target,{get(t,p,r){const local=Reflect.get(t,p,r);return local===undefined?lookup(p):local},has(t,p){return Reflect.has(t,p)||has(p)}})");
        return engine.Invoke(factory, new JsValue[] { native, lookup, has }).AsObject();
    }

    private bool TryWindowFunction(Engine engine, IWindow caller, IWindow target, JsValue property,
        out Realm realm, out Function function) {
        realm = null!;
        function = null!;
        if (!property.IsString()) return false;
        lock (sync) {
            var current = ResolveWindow(target);
            if (!SameScriptOrigin(caller.Document, current.Document))
                throw DomError(engine, "SecurityError", "The target window is not same-origin.");
            if (!_realms.TryGetValue(caller, out var callerRealm) || !ReferenceEquals(callerRealm.Engine, engine) ||
                !_realms.TryGetValue(current, out var targetRealm) || targetRealm.InitialGlobalKeys is null ||
                targetRealm.InitialGlobalKeys.Contains(property.AsString())) return false;
            if (targetRealm.Engine.Global.GetOwnProperty(property).Value is not Function value) return false;
            realm = targetRealm;
            function = value;
            return true;
        }
    }

    private JsValue InvokeWindowFunction(Engine engine, IWindow caller, IWindow target, Realm owner,
        Function function, JsValue receiver, JsValue[] args) {
        lock (sync) {
            RetireDetached();
            var current = ResolveWindow(target);
            if (!SameScriptOrigin(caller.Document, current.Document))
                throw DomError(engine, "SecurityError", "The target window is not same-origin.");
            if (!_realms.TryGetValue(caller, out var callerRealm) || !ReferenceEquals(callerRealm.Engine, engine) ||
                !_realms.TryGetValue(current, out var targetRealm) || !ReferenceEquals(targetRealm, owner))
                throw DomError(engine, "InvalidStateError", "The window function's realm is no longer active.");
            JsValue thisValue = receiver.ToObject() is IWindow self && ReferenceEquals(ResolveWindow(self), current)
                ? owner.Engine.Global : JsValue.Undefined;
            var transferred = args.Select(value => TransferWindowValue(value, engine, caller, owner.Engine, current, engine)).ToArray();
            using var scope = entry.Enter(caller.Document);
            JsValue result;
            try { result = owner.Engine.Invoke(function, thisValue, transferred); }
            catch (JavaScriptException error) {
                string name = error.Error is ObjectInstance errorObject && errorObject.Get("name").IsString()
                    ? errorObject.Get("name").AsString() : "Error";
                throw DomError(engine, name, error.Message);
            }
            return TransferWindowValue(result, owner.Engine, current, engine, caller, engine);
        }
    }

    private JsValue TransferWindowValue(JsValue value, Engine sourceEngine, IWindow sourceWindow,
        Engine destinationEngine, IWindow destinationWindow, Engine errorEngine) {
        if (value.IsUndefined()) return JsValue.Undefined;
        if (value.IsNull()) return JsValue.Null;
        if (value.IsString()) return JsValue.FromObject(destinationEngine, value.AsString());
        if (value.IsNumber()) return JsValue.FromObject(destinationEngine, value.AsNumber());
        if (value.IsBoolean()) return JsValue.FromObject(destinationEngine, value.AsBoolean());
        if (ReferenceEquals(value, sourceEngine.Global)) return WrapWindow(destinationEngine, destinationWindow, sourceWindow);
        object? native = value.ToObject();
        if (native is IWindow window) return WrapWindow(destinationEngine, destinationWindow, window);
        if (native is INode node) return JsValue.FromObject(destinationEngine, node);
        throw DomError(errorEngine, "NotSupportedError", "Cross-realm window calls require primitive or DOM arguments and results.");
    }

    private static bool SameScriptOrigin(IDocument left, IDocument right) {
        string origin = RuntimeDocumentUrls.Origin(left);
        return origin != "null" && string.Equals(origin, RuntimeDocumentUrls.Origin(right), StringComparison.OrdinalIgnoreCase);
    }
}
