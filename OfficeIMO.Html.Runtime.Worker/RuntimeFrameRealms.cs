using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// A session owns every interpreter realm. Child documents may join only through
// the qualified same-origin, script-enabled frame path; all realms share the
// session lock and resource/error accounting while retaining distinct globals.
internal sealed class RuntimeFrameBudget(HtmlScriptRequest options) {
    private readonly object _sync = new();
    private int _childRealms;
    private int _messages;

    internal bool TryReserveChildRealm() {
        lock (_sync) {
            if (_childRealms >= options.MaxChildFrameRealms) return false;
            _childRealms++;
            return true;
        }
    }

    internal bool TryReserveMessage() {
        lock (_sync) {
            if (_messages >= options.MaxFrameMessages) return false;
            _messages++;
            return true;
        }
    }
}

internal sealed class RuntimeFrameRealms(HtmlScriptRequest options, RuntimeFrameBudget budget, object sync, RuntimeScriptErrors errors) {
    private readonly Dictionary<IWindow, Realm> _realms = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<IBrowsingContext, Realm> _contexts = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<IWindow> _reserved = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<IWindow> _blocked = new(ReferenceEqualityComparer.Instance);
    private readonly HashSet<IBrowsingContext> _blockedContexts = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<IWindow, RuntimeModuleLoader> _pendingModules = new(ReferenceEqualityComparer.Instance);
    private readonly List<RetiringRealm> _retiring = [];

    internal bool CanExecute(IDocument document) {
        IWindow? window = document.DefaultView;
        if (window == null) return false;
        IBrowsingContext context = document.Context;
        if (context.Parent == null) return true;
        if ((context.Security & (Sandboxes.Scripts | Sandboxes.Origin)) != 0) return false;
        IWindow? parent = context.Parent.Current;
        if (parent == null || !_realms.ContainsKey(parent)) return false;
        return SameOrigin(parent.Document, document);
    }

    internal bool Reserve(IDocument document) {
        IWindow window = document.DefaultView ?? throw new HtmlScriptRuntimeException("A script realm requires a window.");
        lock (sync) {
            if (_realms.ContainsKey(window) || _reserved.Contains(window) || _blocked.Contains(window)) return false;
            if (_blockedContexts.Contains(document.Context)) {
                _blocked.Add(window);
                return false;
            }
            if (_contexts.TryGetValue(document.Context, out Realm? prior) && !ReferenceEquals(prior.Window, window)) {
                Retire(prior, blockContext: true, RetirementKind.Replacement);
                _blocked.Add(window);
                return false;
            }
            if (!CanExecute(document)) return false;
            if (document.Context.Parent != null) {
                if (!budget.TryReserveChildRealm()) {
                    _blocked.Add(window);
                    return false;
                }
            }
            _reserved.Add(window);
            return true;
        }
    }

    internal void Register(IDocument document, Engine engine, RuntimeEventLoop loop) {
        IWindow window = document.DefaultView ?? throw new HtmlScriptRuntimeException("A script realm requires a window.");
        lock (sync) {
            if (!_pendingModules.Remove(window, out RuntimeModuleLoader? modules))
                throw new HtmlScriptRuntimeException("The script realm does not have a module loader.");
            _reserved.Remove(window);
            ObjectInstance cloneTransport = RuntimeStructuredClone.CreateTransport(engine, options.MaxFrameMessageCharacters);
            var dispatch = new ClrFunction(engine, "dispatchMessageEvent", (_, args) => {
                var message = (AngleSharp.Dom.Events.Event)args[0].ToObject()!;
                RuntimeEventTrust.Set(message, true);
                window.Dispatch(message);
                return JsValue.Undefined;
            });
            var realm = new Realm(window, document.Context, engine, loop,
                cloneTransport.Get("decode"),
                engine.Invoke(engine.Evaluate("dispatch=>(data,origin,source)=>{const event=new MessageEvent('message');Object.defineProperties(event,{data:{value:data,enumerable:true},origin:{value:origin,enumerable:true},source:{value:source,enumerable:true},ports:{value:Object.freeze([]),enumerable:true}});dispatch(event)}"), new JsValue[] { dispatch }),
                modules, new CancellationTokenSource());
            _realms[window] = realm;
            _contexts[document.Context] = realm;
        }
    }

    internal RuntimeModuleLoader AttachModules(IDocument document, RuntimeModuleLoader modules) {
        IWindow window = document.DefaultView ?? throw new HtmlScriptRuntimeException("A module loader requires a window.");
        lock (sync) {
            if (!_reserved.Contains(window) || _pendingModules.ContainsKey(window) || _realms.ContainsKey(window))
                throw new HtmlScriptRuntimeException("The script realm module loader was created out of sequence.");
            _pendingModules.Add(window, modules);
            return modules;
        }
    }

    internal RuntimeModuleLoader? ModulesFor(IBrowsingContext context) {
        lock (sync) return context.Current != null && _realms.TryGetValue(context.Current, out Realm? realm)
            ? realm.Modules
            : null;
    }

    internal CancellationToken LifetimeFor(IBrowsingContext context) {
        lock (sync) return _contexts.TryGetValue(context, out Realm? realm)
            ? realm.Lifetime.Token
            : new CancellationToken(canceled: true);
    }

    internal void Own(IDocument document, IDisposable resource) {
        lock (sync) {
            if (_contexts.TryGetValue(document.Context, out Realm? realm)
                && ReferenceEquals(realm.Window, document.DefaultView)) {
                realm.Resources.Add(resource);
                return;
            }
        }
        resource.Dispose();
    }

    internal void BeginFrameNavigation(IElement host) {
        if (host is not IHtmlInlineFrameElement frame) return;
        lock (sync) {
            IBrowsingContext? context = frame.ContentDocument?.Context;
            Realm? realm = context != null && _contexts.TryGetValue(context, out Realm? current) ? current : null;
            if (realm != null) Retire(realm, blockContext: true, RetirementKind.Replacement);
        }
    }

    internal Engine? EngineFor(IBrowsingContext context) {
        lock (sync) return context.Current != null && _realms.TryGetValue(context.Current, out Realm? realm) ? realm.Engine : null;
    }

    internal Engine? EngineFor(IElement element) {
        lock (sync) {
            if (element.Owner == null) return null;
            if (element.Owner.DefaultView is { } window && _realms.TryGetValue(window, out Realm? direct)) return direct.Engine;
            return element.Owner.Context.Current != null && _realms.TryGetValue(element.Owner.Context.Current, out Realm? current)
                ? current.Engine : null;
        }
    }

    internal IReadOnlyList<Engine> Engines() {
        lock (sync) return _realms.Values.Select(realm => realm.Engine).ToArray();
    }

    internal void RetireDetached() {
        lock (sync) {
            Realm? root = _contexts.Values.FirstOrDefault(realm => realm.Context.Parent == null);
            if (root == null) return;
            var attached = new HashSet<IBrowsingContext>(ReferenceEqualityComparer.Instance);
            CollectAttachedContexts(root.Window.Document, attached);
            foreach (Realm realm in _contexts.Values
                         .Where(realm => realm.Context.Parent != null
                             && (!attached.Contains(realm.Context) || !ReferenceEquals(realm.Context.Current, realm.Window)))
                         .ToArray()) {
                RetirementKind kind = attached.Contains(realm.Context)
                    ? RetirementKind.Replacement
                    : RetirementKind.Detached;
                Retire(realm, blockContext: true, kind);
            }
            DisposeSettledRetirements();
        }
    }

    internal void DisposeAll() {
        lock (sync) {
            foreach (Realm realm in _realms.Values.ToArray()) Retire(realm, blockContext: false, RetirementKind.Detached, force: true);
            foreach (RetiringRealm retirement in _retiring) retirement.Realm.Loop.Dispose();
            _retiring.Clear();
            _reserved.Clear();
            _pendingModules.Clear();
            _blocked.Clear();
            _blockedContexts.Clear();
        }
    }

    internal void InstallMessaging(Engine engine, IWindow source) {
        JsValue encode = RuntimeStructuredClone.CreateTransport(engine, options.MaxFrameMessageCharacters).Get("encode");
        JsValue hasTransfers = engine.Evaluate("value=>Array.isArray(value)&&value.length!==0");
        var post = new ClrFunction(engine, "postMessage", (receiver, args) => {
            IWindow? target = ReferenceEquals(receiver, engine.Global) ? source : receiver.ToObject() as IWindow;
            if (target == null) throw TypeError(engine, "postMessage requires a Window receiver.");
            string targetOrigin = args.Length > 1 && !args[1].IsUndefined() ? TypeConverter.ToString(args[1]) : "/";
            if (args.Length > 2 && !args[2].IsUndefined() && engine.Invoke(hasTransfers, args[2]).AsBoolean())
                throw DomError(engine, "DataCloneError", "Transfer lists are not supported.");
            JsValue serialized;
            try { serialized = engine.Invoke(encode, args.Length == 0 ? JsValue.Undefined : args[0]); }
            catch (JavaScriptException error) when (error.Message.Contains("frame message exceeds its character budget", StringComparison.OrdinalIgnoreCase)) {
                throw DomError(engine, "QuotaExceededError", error.Message);
            }
            catch (JavaScriptException) { throw; }
            catch (Exception error) { throw DomError(engine, "DataCloneError", error.Message); }
            if (!serialized.IsString()) throw DomError(engine, "DataCloneError", "The message could not be cloned.");
            Post(engine, source, target, targetOrigin, serialized.AsString());
            return JsValue.Undefined;
        });
        var descriptor = new PropertyDescriptor(post, true, false, true);
        engine.Global.FastSetProperty("postMessage", descriptor);
        if (engine.Global.Get("Window") is Function constructor && constructor.Get("prototype") is ObjectInstance prototype)
            prototype.FastSetProperty("postMessage", descriptor);
    }

    private void Post(Engine sourceEngine, IWindow source, IWindow target, string targetOrigin, string json) {
        Realm targetRealm;
        string sourceOrigin = Origin(source.Document);
        lock (sync) {
            if (!_realms.TryGetValue(target, out targetRealm!)) {
                targetRealm = _realms.Values.FirstOrDefault(candidate =>
                    ReferenceEquals(candidate.Window.Document, target.Document)
                    || ReferenceEquals(candidate.Window.Document.Context, target.Document.Context))!;
                if (targetRealm == null) return;
            }
            string actualTargetOrigin = Origin(target.Document);
            if (!AcceptsTargetOrigin(sourceEngine, targetOrigin, sourceOrigin, actualTargetOrigin)) return;
            if (!budget.TryReserveMessage()) {
                throw DomError(sourceEngine, "QuotaExceededError", "Parent/child frame message budget exceeded.");
            }
        }
        targetRealm.Loop.Enqueue(_ => {
            lock (sync) {
                if (!_realms.ContainsKey(source) || !_realms.TryGetValue(targetRealm.Window, out Realm? currentTarget)
                    || !ReferenceEquals(currentTarget, targetRealm)) return;
                try {
                    JsValue data = targetRealm.Engine.Invoke(targetRealm.DecodeMessage, json);
                    targetRealm.Engine.Invoke(targetRealm.DispatchMessage, new JsValue[] {
                        data, sourceOrigin, JsValue.FromObject(targetRealm.Engine, source)
                    });
                } catch (Exception error) { errors.Report(error.Message); }
            }
        }, TaskPriority.Normal);
    }

    private static bool AcceptsTargetOrigin(Engine engine, string requested, string sourceOrigin, string targetOrigin) {
        if (requested == "*") return true;
        if (requested == "/") return string.Equals(sourceOrigin, targetOrigin, StringComparison.OrdinalIgnoreCase);
        if (!Uri.TryCreate(requested, UriKind.Absolute, out Uri? uri)
            || uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps
            || uri.UserInfo.Length != 0)
            throw DomError(engine, "SyntaxError", "The target origin must be '*', '/', or an absolute HTTP(S) origin.");
        return string.Equals(uri.GetLeftPart(UriPartial.Authority), targetOrigin, StringComparison.OrdinalIgnoreCase);
    }

    private static bool SameOrigin(IDocument left, IDocument right) =>
        string.Equals(Origin(left), Origin(right), StringComparison.OrdinalIgnoreCase);

    private static string Origin(IDocument document) =>
        RuntimeDocumentUrls.Origin(document);

    private static void CollectAttachedContexts(IDocument document, HashSet<IBrowsingContext> attached) {
        foreach (IHtmlInlineFrameElement frame in document.QuerySelectorAll("iframe").OfType<IHtmlInlineFrameElement>()) {
            if (frame.ContentDocument is not { } child || !attached.Add(child.Context)) continue;
            CollectAttachedContexts(child, attached);
        }
    }

    private void Retire(Realm realm, bool blockContext, RetirementKind kind, bool force = false) {
        _realms.Remove(realm.Window);
        _contexts.Remove(realm.Context);
        _reserved.Remove(realm.Window);
        _blocked.Add(realm.Window);
        if (blockContext) _blockedContexts.Add(realm.Context);
        foreach (IDisposable resource in realm.Resources) resource.Dispose();
        realm.Resources.Clear();
        realm.Lifetime.Cancel();
        realm.Lifetime.Dispose();
        realm.Window.Dispose();
        if (force || CanDispose(realm, kind)) realm.Loop.Dispose();
        else _retiring.Add(new RetiringRealm(realm, kind));
    }

    private void DisposeSettledRetirements() {
        for (int index = _retiring.Count - 1; index >= 0; index--) {
            RetiringRealm retirement = _retiring[index];
            if (!CanDispose(retirement.Realm, retirement.Kind)) continue;
            retirement.Realm.Loop.Dispose();
            _retiring.RemoveAt(index);
        }
    }

    private static bool CanDispose(Realm realm, RetirementKind kind) {
        if (kind == RetirementKind.Detached) return realm.Window.Document.ReadyState == DocumentReadyState.Complete;
        IWindow? current = realm.Context.Current;
        return current != null
            && !ReferenceEquals(current, realm.Window)
            && current.Document.ReadyState == DocumentReadyState.Complete;
    }

    private static JavaScriptException TypeError(Engine engine, string message) =>
        new(engine.Intrinsics.TypeError, message);

    private static JavaScriptException DomError(Engine engine, string name, string message) {
        var error = engine.Intrinsics.Error.Construct(message);
        error.Set("name", name);
        return new JavaScriptException(error);
    }

    private sealed record Realm(IWindow Window, IBrowsingContext Context, Engine Engine, RuntimeEventLoop Loop,
        JsValue DecodeMessage, JsValue DispatchMessage, RuntimeModuleLoader Modules, CancellationTokenSource Lifetime) {
        internal List<IDisposable> Resources { get; } = [];
    }

    private sealed record RetiringRealm(Realm Realm, RetirementKind Kind);
    private enum RetirementKind { Detached, Replacement }
}
