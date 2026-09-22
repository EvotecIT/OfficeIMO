using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using AngleSharp.Js;
using AngleSharp.Scripting;
using Jint;
using AngleSharp.Html.Dom.Events;
using AngleSharp.Html.Dom;
using AngleSharp.Html;
using AngleSharp.Html.Parser;
using AngleSharp.Io;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class ScriptedDocumentSession : IDisposable {
    private IBrowsingContext _context;
    private readonly IConfiguration _configuration;
    private readonly RuntimeScriptErrors _errors;
    private readonly RuntimeResourceLoader _resources;
    private readonly RuntimeSubresourceIntegrity _integrity;
    private readonly RuntimeModuleSourceCache _moduleSources;
    private Engine _engine = null!;
    private readonly object _realmSync = new();
    private readonly RuntimeScriptEntry _scriptEntry = new();
    private readonly RuntimeFrameRealms _realms;
    private readonly RuntimeAuxiliaryWindows _auxiliary;
    private readonly JsScriptingService _providerScripting;
    private readonly HtmlScriptRequest _options;
    private IDocument _document = null!;
    private IEventLoop _loop = null!;
    private readonly RuntimeFocusController _focus = new();
    private RuntimeAutomation _automation = null!;
    private CancellationToken _activeCommandToken;
    private readonly RuntimeScriptingService _scripting;
    private RuntimeHistoryBindings _history = null!;
    private readonly ConditionalWeakTable<IDocument, RuntimeHistoryBindings> _histories = new();
    private readonly Func<long> _currentRevision;
    private readonly Action _markRevision;
    private readonly RuntimeDiagnostics _diagnostics;
    private readonly RuntimeBrowsingStorage _storage;
    private readonly RuntimeBrowsingHistory _historyState;
    private Action<RuntimeNavigation>? _navigate;

    private ScriptedDocumentSession(HtmlScriptRequest options, RuntimeResourceBudget budget, RuntimeFrameBudget frameBudget, RuntimeBrowsingStorage storage,
        RuntimeBrowsingHistory history, RuntimeDiagnostics diagnostics, Action<RuntimeNavigation>? navigate, Func<long> currentRevision, Action markRevision) {
        _options = options;
        _diagnostics = diagnostics;
        _storage = storage;
        _historyState = history;
        _navigate = navigate;
        _errors = new RuntimeScriptErrors(options.MaxPendingPromiseRejections, diagnostics);
        _realms = new RuntimeFrameRealms(options, frameBudget, _realmSync, _errors, _scriptEntry);
        _auxiliary = new RuntimeAuxiliaryWindows(options, frameBudget, _realms, _scriptEntry, EnsureEngine);
        _resources = new RuntimeResourceLoader(options, budget, diagnostics);
        _integrity = new RuntimeSubresourceIntegrity(options.MaxModuleIntegrityMetadataCharacters);
        _moduleSources = new RuntimeModuleSourceCache(_resources, options.MaxModuleCount,
            options.MaxModuleIntegrityMetadataCharacters, _integrity);
        _currentRevision = currentRevision;
        _markRevision = markRevision;
        var linkRelations = new DefaultLinkRelationFactory();
        linkRelations.Register("modulepreload", link => new RuntimeModulePreloadLinkRelation(link, _moduleSources,
            () => link.Owner == null ? null : _realms.ModulesFor(link.Owner.Context)?.ImportMap,
            () => link.Owner == null ? new CancellationToken(canceled: true) : _realms.LifetimeFor(link.Owner.Context),
            _errors.Report));
        var configuration = Configuration.Default.WithCss().WithJs(new JsScriptingOptions {
                MaxCallStackDepth = 512,
                ConfigureEngine = (window, engineOptions) => {
                    IDocument document = window.Document;
                    RuntimeModuleLoader modules = _realms.AttachModules(document, new RuntimeModuleLoader(document, _moduleSources,
                        () => document.Context.GetService<IEventLoop>() as RuntimeEventLoop
                            ?? throw new HtmlScriptRuntimeException("The module realm does not have an event loop."),
                        () => _realms.EngineFor(document.Context) ?? throw new HtmlScriptRuntimeException("The module realm is no longer active."),
                        () => _realms.LifetimeFor(document.Context)));
                    engineOptions.EnableModules(modules).UseHostFactory(engine => new RuntimeModuleHost(engine, modules, () => _realms.CanExecuteJob(window)));
                }
            })
            .WithEventLoop(context => new RuntimeEventLoop(context, () => _realms.EngineFor(context), _errors, _realmSync, _realms.RetireDetached, _scriptEntry))
            .With(new RuntimeDomRevisionListener(markRevision))
            .With(new RuntimeDomSynchronization(() => _realmSync))
            .With(new RuntimeScriptBlockingStyleSheetEvaluator(options))
            .WithOnly<IIntegrityProvider>(_integrity)
            .WithOnly<ILinkRelationFactory>(linkRelations)
            .Without<AngleSharp.Css.IPseudoClassSelectorFactory>()
            .With((AngleSharp.Css.IPseudoClassSelectorFactory)_focus.CreateSelectors())
            .With<IRequester>(context => new RuntimeResourceRequester(_resources, _errors, () => _realms.ResourceLifetimeFor(context)))
            .WithDefaultLoader(new LoaderOptions { IsResourceLoadingEnabled = true, IsNavigationDisabled = true })
            .WithOnly<IResourceLoader>(context => new RuntimeDocumentResourceLoader(
                context, _moduleSources, () => _realms.ModulesFor(context)?.ImportMap,
                () => _realms.LifetimeFor(context), options));
        _providerScripting = configuration.Services.OfType<JsScriptingService>().Single();
        _scripting = new RuntimeScriptingService(_providerScripting,
            document => _realms.ModulesFor(document.Context),
            document => _realms.LifetimeFor(document.Context),
            EnsureEngine, _realmSync, options, _errors.Report, _scriptEntry);
        configuration = configuration.Without<IScriptingService>()
            .With(_scripting);
        var scriptObservers = configuration.Services.OfType<IAttributeObserver>().Where(observer => observer.GetType().Assembly == typeof(JsScriptingService).Assembly).ToArray();
        // Replace only the script observer; retain CSS and native DOM attribute observers.
        configuration = configuration.Without(scriptObservers)
            .With(new RuntimeEventAttributeObserver(_realms.EngineFor));
        // Local frame loads can complete inside the native attribute observer.
        // Retire the previous realm before that observer can create its replacement.
        configuration = new Configuration(configuration.Services.Prepend(new RuntimeFrameNavigationObserver(_realms)));
        _configuration = configuration;
        _context = CreateRootContext();
    }

    private IBrowsingContext CreateRootContext() {
        var context = BrowsingContext.New(_configuration);
        _loop = context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
        context.GetService<IHtmlParser>()!.Parsing += (_, args) => {
            var document = ((HtmlParseEvent)args).Document;
            EnsureEngine(document);
        };
        return context;
    }

    private Engine? EnsureEngine(IDocument document) {
        lock (_realmSync) {
            if (_realms.EngineFor(document.Context) is { } existing) return existing;
            // A scriptless local parent still needs a realm before a nested
            // child's scripts can use its window. Initialize ancestors root-first.
            var ancestors = new Stack<IDocument>();
            for (var parent = document.Context.Parent?.Active;
                 parent != null && _realms.EngineFor(parent.Context) == null;
                 parent = parent.Context.Parent?.Active) {
                if (ancestors.Count >= _options.MaxChildFrameRealms) return null;
                ancestors.Push(parent);
            }
            while (ancestors.TryPop(out var parent)) {
                if (EnsureEngine(parent) == null) return null;
            }
            if (!_realms.Reserve(document)) return _realms.EngineFor(document.Context);
            Engine engine = _providerScripting.GetOrCreateJint(document);
            var loop = document.Context.GetService<IEventLoop>() as RuntimeEventLoop
                ?? throw new HtmlScriptRuntimeException("The provider did not create a frame event loop.");
            _realms.Register(document, engine, loop);
            bool root = document.Context.Parent == null;
            IDisposable contextErrors = AttachErrors(document.Context);
            IDisposable windowErrors = AttachErrors(document.DefaultView!);
            _realms.Own(document, contextErrors);
            _realms.Own(document, windowErrors);
            _errors.Attach(engine);
            loop.InitializeMicrotasks(engine);
            _scripting.Initialize(engine);
            if (root) {
                _engine = engine;
                _loop = loop;
            }
            var normalizeWindow = RuntimeWindowBindings.Install(engine, document.DefaultView!, _realms);
            _auxiliary.Install(engine, document.DefaultView!);
            RuntimeNativeNavigationBindings.Install(engine, document.DefaultView!, _realms.IsAuxiliaryNavigation);
            RuntimeEventBindings.Install(engine, document.DefaultView!, _errors.Report, normalizeWindow, resource => _realms.Own(document, resource));
            RuntimeDocumentOpenBindings.Install(engine, document, _scriptEntry, opened => {
                if (_histories.TryGetValue(opened, out var history)) history.RewriteDocumentUrl();
            }, (window, args) => _auxiliary.Open(engine, window, args));
            RuntimeUrlBindings.Install(engine, document.DefaultView!);
            RuntimeConsoleBindings.Install(engine, _diagnostics);
            RuntimeObserverBindings.Install(engine, document, _errors.Report, loop.EnqueueMicrotask, resource => _realms.Own(document, resource));
            RuntimeStorageBindings.Install(engine, _options.MaxStorageCharacters, _storage,
                RuntimeDocumentUrls.Origin(document));
            var fetch = new RuntimeFetchBindings(engine, document, loop, _resources, _options, _errors);
            _realms.Own(document, fetch);
            RuntimeWindowNavigationBindings.Install(engine, document.DefaultView!, _realms);
            if (!root) {
                if (_realms.IsAuxiliary(document.DefaultView!)) {
                    var popupViewport = new RuntimeViewport((IHtmlDocument)document, _options, () => _activeCommandToken, _resources.Capture);
                    _histories.Add(document, new RuntimeHistoryBindings(engine, document, loop, _options, popupViewport, installLocation: false));
                }
                _realms.CaptureWindowSurface(document);
                return engine;
            }
            var viewport = new RuntimeViewport((IHtmlDocument)document, _options, () => _activeCommandToken, _resources.Capture);
            _history = new RuntimeHistoryBindings(engine, document, loop, _options, viewport, _historyState, _navigate);
            _histories.Add(document, _history);
            _automation = new RuntimeAutomation(document, _options, _focus, _history, viewport, engine, _diagnostics);
            RuntimeInteractionBindings.Install(engine, document, _focus, _automation, _options.DevicePixelRatio);
            RuntimeSelectBindings.Install(engine);
            _realms.CaptureWindowSurface(document);
            return engine;
        }
    }

    internal static async Task<ScriptedDocumentSession> OpenAsync(HtmlScriptRequest request, RuntimeResourceBudget budget, RuntimeFrameBudget frameBudget,
        RuntimeBrowsingStorage storage, RuntimeBrowsingHistory history, RuntimeDiagnostics diagnostics, Action<RuntimeNavigation>? navigate,
        Func<long> currentRevision, Action markRevision, CancellationToken token, HtmlRuntimeResource? source = null) {
        var session = new ScriptedDocumentSession(request, budget, frameBudget, storage, history, diagnostics, navigate, currentRevision, markRevision);
        try {
            await session.LoadDocumentAsync(request, token, source);
            return session;
        } catch { session.Dispose(); throw; }
    }

    // The browsing session retains this owner across navigation. Only the root
    // document and its frame descendants retire; auxiliary realms keep running.
    internal async Task ReplaceDocumentAsync(HtmlScriptRequest request, Action<RuntimeNavigation>? navigate,
        CancellationToken token, HtmlRuntimeResource? source) {
        lock (_realmSync) {
            foreach (Engine engine in _realms.RetireRoot()) _scripting.Retire(engine);
            _context.Dispose();
            _focus.Reset();
            _navigate = navigate;
            _context = CreateRootContext();
        }
        await LoadDocumentAsync(request, token, source);
    }

    private async Task LoadDocumentAsync(HtmlScriptRequest request, CancellationToken token, HtmlRuntimeResource? source) {
            string html = source == null ? request.Html : RuntimeHtmlNavigationSource.Decode(source, request.MaxInputCharacters);
            _document = await _context.OpenAsync(response => {
                response.Address(request.DocumentUrl);
                response.Content(html);
                if (source != null) response.Header("Content-Type", source.ContentType).Status(source.StatusCode);
            }, token).WaitUntilAvailable(token);
            await _scripting.WaitForModuleEvaluationsAsync(_document, token);
            _loop = _context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
            foreach (string script in request.Scripts) await ExecuteAsync(script, token);
            await OnLoop(() => true, token);
    }

    internal Task ExecuteAsync(string script, CancellationToken token) => OnLoop(() => { _scripting.EvaluateScript(_document, script, "text/javascript", RuntimeDocumentUrls.Base(_document)); return true; }, token);
    internal Task NavigateAsync(string target, bool replace, CancellationToken token) => OnLoop(() => { _history.NavigateFragment(target, replace); return true; }, token);
    internal Task RestoreTraversalAsync(bool dispatchPopState, CancellationToken token) => OnLoop(() => { _history.RestoreTraversal(dispatchPopState); return true; }, token);
    internal Task ReloadAsync(CancellationToken token) => OnLoop(() => { _history.Reload(); return true; }, token);
    internal Uri DocumentUrl => new(_document.Url);
    internal Task<bool> PromptToUnloadAsync(CancellationToken token) => OnLoop(() => {
        using var unload = ((Document)_document).BeginUnloadScope();
        var beforeUnload = new AngleSharp.Js.Dom.BeforeUnloadEvent();
        RuntimeEventTrust.Set(beforeUnload, true);
        _document.DefaultView!.Dispatch(beforeUnload);
        return !beforeUnload.IsDefaultPrevented && beforeUnload.ReturnValue.Length == 0;
    }, token);
    internal Task CommitUnloadAsync(CancellationToken token) => OnLoop(() => {
        using var unload = ((Document)_document).BeginUnloadScope();
        _document.DefaultView!.Dispatch(new PageTransitionEvent("pagehide", bubbles: false, cancelable: false, persisted: false));
        _document.DefaultView.Dispatch(new Event("unload", bubbles: false, cancelable: false));
        return true;
    }, token);

    internal Task<HtmlAutomationResult> AutomateAsync(HtmlAutomationRequest request, string pageId, long revision, CancellationToken token) =>
        OnLoop(() => _automation.Run(request, pageId, revision, token), token);

    internal Task<HtmlPageObservation> ObserveAsync(HtmlPageObservationRequest request, string contextId, string pageId, CancellationToken token) =>
        OnLoop(() => _automation.Observe(request, contextId, pageId, _currentRevision(), token), token);

    internal Task<string> EvaluateAsync(string expression, CancellationToken token) => OnLoop(() => {
        var value = _engine.Evaluate("JSON.stringify((" + expression + "\n))", RuntimeDocumentUrls.Base(_document));
        if (!value.IsString()) throw new HtmlScriptRuntimeException("The expression did not produce a JSON value.");
        return value.AsString();
    }, token);

    internal Task<(bool Ready, HtmlRuntimeWireDocument? Document)> ProbeAsync(string expression, bool capture, CancellationToken token) => OnLoop(() => {
                if (_scripting.EvaluateScript(_document, expression, "text/javascript", RuntimeDocumentUrls.Base(_document)) is not Jint.Native.JsValue ready || !ready.IsBoolean() || !ready.AsBoolean()) return (Ready: false, Document: (HtmlRuntimeWireDocument?)null);
                _errors.ThrowIfFailed();
                // Readiness and capture share a task so timers cannot mutate between them.
                var document = capture ? RuntimeDomCapture.Capture(_document, _options, token) : null;
                if (document != null) { document.DocumentUrl = new Uri(_document.Url); document.BaseUri = new Uri(RuntimeDocumentUrls.Base(_document)); document.Resources = _resources.Capture().ToList(); }
                return (Ready: true, Document: document);
            }, token);

    private Task<T> OnLoop<T>(Func<T> action, CancellationToken token) {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _loop.Enqueue(_ => {
            lock (_realmSync) {
                CancellationToken previousToken = _activeCommandToken;
                _activeCommandToken = token;
                try {
                    token.ThrowIfCancellationRequested();
                    // Native timer callbacks may leave a rejection whose catch is already
                    // queued. Preserve fatal errors, but check rejections after the checkpoint.
                    _errors.ThrowIfFailed(includeRejections: false);
                    ProcessRealmTasks();
                    _errors.ThrowIfFailed();
                    T result = action();
                    // Native DOM actions can invoke JS callbacks without entering Evaluate.
                    // Complete their promise jobs before admitting the next session command.
                    ProcessRealmTasks();
                    _errors.ThrowIfFailed();
                    completion.TrySetResult(result);
                } catch (Exception error) {
                    // Preserve a listener's reported JS error when the provider rethrows an
                    // outer invocation failure with an empty or less useful message.
                    try { _errors.ThrowIfFailed(); }
                    catch (Exception tracked) { error = tracked; }
                    completion.TrySetException(error);
                } finally {
                    _activeCommandToken = previousToken;
                }
            }
        }, TaskPriority.Normal);
        return completion.Task.WaitAsync(token);
    }

    public void Dispose() {
        lock (_realmSync) DisposeCore();
    }

    private void ProcessRealmTasks() {
        _realms.RetireDetached();
        foreach (Engine engine in _realms.Engines()) engine.Advanced.ProcessTasks();
    }

    private IDisposable AttachErrors(IEventTarget target) {
        DomEventHandler handler = (_, error) => _errors.Report(error switch {
            AngleSharp.Dom.Events.ErrorEvent scriptError => scriptError.Message,
            AngleSharp.Browser.Dom.Events.TrackEvent tracked => tracked.Error?.GetBaseException().Message ?? "Script execution failed.",
            _ => "Script execution failed."
        });
        target.AddEventListener("error", handler);
        return new EventSubscription(target, handler);
    }

    private void DisposeCore() {
        _scripting.Dispose();
        _realms.DisposeAll();
        _resources.Dispose();
        _context.Dispose();
    }

    private sealed class EventSubscription(IEventTarget target, DomEventHandler handler) : IDisposable {
        private IEventTarget? _target = target;

        public void Dispose() {
            IEventTarget? current = Interlocked.Exchange(ref _target, null);
            current?.RemoveEventListener("error", handler);
        }
    }
}
