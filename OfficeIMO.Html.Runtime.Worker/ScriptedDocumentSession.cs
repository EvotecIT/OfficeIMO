using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Js;
using AngleSharp.Scripting;
using Jint;
using AngleSharp.Html.Dom.Events;
using AngleSharp.Html.Parser;
using AngleSharp.Io;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class ScriptedDocumentSession : IDisposable {
    private readonly IBrowsingContext _context;
    private readonly RuntimeScriptErrors _errors;
    private readonly RuntimeResourceLoader _resources;
    private Engine _engine = null!;
    private readonly HtmlScriptRequest _options;
    private IDocument _document = null!;
    private IEventLoop _loop = null!;
    private RuntimeFetchBindings? _fetch;
    private readonly RuntimeFocusController _focus = new();
    private RuntimeAutomation _automation = null!;
    private RuntimeModuleLoader _modules = null!;
    private readonly RuntimeScriptingService _scripting;
    private RuntimeHistoryBindings _history = null!;

    private ScriptedDocumentSession(HtmlScriptRequest options) {
        _options = options;
        _errors = new RuntimeScriptErrors(options.MaxPendingPromiseRejections);
        _resources = new RuntimeResourceLoader(options);
        var configuration = Configuration.Default.WithCss().WithJs(new JsScriptingOptions {
                MaxCallStackDepth = 512,
                ConfigureEngine = (window, engineOptions) => {
                    if (_modules != null) throw new HtmlScriptRuntimeException("Additional worker or window interpreters are outside this session profile.");
                    _modules = new RuntimeModuleLoader(window.Document, _resources, () => _loop, () => _engine, options.MaxModuleCount);
                    engineOptions.EnableModules(_modules).UseHostFactory(_ => new RuntimeModuleHost());
                }
            })
            .WithEventLoop(context => new RuntimeEventLoop(context, () => _engine, _errors))
            .With(new RuntimeDocumentUrls.MutationListener())
            .With(new RuntimeDomSynchronization(() => _engine))
            .Without<AngleSharp.Css.IPseudoClassSelectorFactory>()
            .With((AngleSharp.Css.IPseudoClassSelectorFactory)_focus.CreateSelectors())
            .With(new RuntimeResourceRequester(_resources, _errors))
            .WithDefaultLoader(new LoaderOptions { IsResourceLoadingEnabled = true, IsNavigationDisabled = true })
            .WithOnly<IResourceLoader>(context => new RuntimeDocumentResourceLoader(context));
        var scripting = configuration.Services.OfType<JsScriptingService>().Single();
        _scripting = new RuntimeScriptingService(scripting, () => _modules, options, _errors.Report);
        configuration = configuration.Without<IScriptingService>()
            .With(scripting).With(_scripting);
        var scriptObservers = configuration.Services.OfType<IAttributeObserver>().Where(observer => observer.GetType().Assembly == typeof(JsScriptingService).Assembly).ToArray();
        // Replace only the script observer; retain CSS and native DOM attribute observers.
        configuration = configuration.Without(scriptObservers).With(new RuntimeEventAttributeObserver(host => _engine ?? scripting.GetOrCreateJint(host.Owner!)));
        _context = BrowsingContext.New(configuration);
        _loop = _context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
        _context.AddEventListener("error", (_, error) => _errors.Report(error switch {
            AngleSharp.Dom.Events.ErrorEvent scriptError => scriptError.Message,
            AngleSharp.Browser.Dom.Events.TrackEvent tracked => tracked.Error?.Message ?? "Script execution failed.",
            _ => "Script execution failed."
        }));
        _context.GetService<IHtmlParser>()!.Parsing += (_, args) => {
            // Fragment parsing (for example innerHTML) must not replace the live engine
            // or wrap its prototypes again and lose existing listener registrations.
            if (_engine != null) return;
            var document = ((HtmlParseEvent)args).Document;
            _engine = _context.GetService<JsScriptingService>()!.GetOrCreateJint(document);
            _errors.Attach(_engine);
            ((RuntimeEventLoop)_loop).InitializeMicrotasks(_engine);
            _scripting.Initialize(_engine);
            var normalizeWindow = RuntimeWindowBindings.Install(_engine, document.DefaultView!);
            RuntimeEventBindings.Install(_engine, document.DefaultView!, _errors.Report, normalizeWindow);
            RuntimeUrlBindings.Install(_engine, document.DefaultView!);
            RuntimeObserverBindings.Install(_engine, document, _errors.Report, ((RuntimeEventLoop)_loop).EnqueueMicrotask);
            RuntimeStorageBindings.Install(_engine, options.MaxStorageCharacters);
            _history = new RuntimeHistoryBindings(_engine, document, _loop, options);
            _automation = new RuntimeAutomation(document, options, _focus, _history);
            RuntimeInteractionBindings.Install(_engine, document, _focus, _automation);
            RuntimeSelectBindings.Install(_engine);
            _fetch = new RuntimeFetchBindings(_engine, document, _loop, _resources, options, _errors);
        };
    }

    internal static async Task<ScriptedDocumentSession> OpenAsync(HtmlScriptRequest request, CancellationToken token) {
        var session = new ScriptedDocumentSession(request);
        try {
            session._document = await session._context.OpenAsync(source => source.Address(request.DocumentUrl.AbsoluteUri).Content(request.Html), token).WaitUntilAvailable(token);
            await session._scripting.WaitForModuleEvaluationsAsync(token);
            session._loop = session._context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
            foreach (string script in request.Scripts) await session.ExecuteAsync(script, token);
            await session.OnLoop(() => true, token);
            return session;
        } catch { session.Dispose(); throw; }
    }

    internal Task ExecuteAsync(string script, CancellationToken token) => OnLoop(() => { _scripting.EvaluateScript(_document, script, "text/javascript", RuntimeDocumentUrls.Base(_document)); return true; }, token);

    internal async Task<HtmlAutomationResult> AutomateAsync(HtmlAutomationRequest request, CancellationToken token) {
        while (true) {
            token.ThrowIfCancellationRequested();
            HtmlAutomationResult result = await OnLoop(() => _automation.Run(request, token), token);
            if (!request.WaitForReady || result.Status is not (HtmlAutomationStatus.NotFound or HtmlAutomationStatus.NotReady)) return result;
            await Task.Delay(_options.PollInterval, token);
        }
    }

    internal Task<string> EvaluateAsync(string expression, CancellationToken token) => OnLoop(() => {
        var value = _engine.Evaluate("JSON.stringify((" + expression + "\n))", RuntimeDocumentUrls.Base(_document));
        if (!value.IsString()) throw new HtmlScriptRuntimeException("The expression did not produce a JSON value.");
        return value.AsString();
    }, token);

    internal async Task<HtmlRuntimeWireDocument?> WaitAsync(string expression, bool capture, CancellationToken token) {
        while (true) {
            token.ThrowIfCancellationRequested();
            var result = await OnLoop(() => {
                if (_scripting.EvaluateScript(_document, expression, "text/javascript", RuntimeDocumentUrls.Base(_document)) is not Jint.Native.JsValue ready || !ready.IsBoolean() || !ready.AsBoolean()) return (Ready: false, Document: (HtmlRuntimeWireDocument?)null);
                _errors.ThrowIfFailed();
                // Readiness and capture share a task so timers cannot mutate between them.
                var document = capture ? RuntimeDomCapture.Capture(_document, _options, token) : null;
                if (document != null) { document.DocumentUrl = new Uri(_document.Url); document.BaseUri = new Uri(RuntimeDocumentUrls.Base(_document)); document.Resources = _resources.Capture().ToList(); }
                return (Ready: true, Document: document);
            }, token);
            if (result.Ready) return result.Document;
            await Task.Delay(_options.PollInterval, token);
        }
    }

    private Task<T> OnLoop<T>(Func<T> action, CancellationToken token) {
        var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _loop.Enqueue(_ => {
            lock (_engine) {
                try {
                    token.ThrowIfCancellationRequested();
                    // Native timer callbacks may leave a rejection whose catch is already
                    // queued. Preserve fatal errors, but check rejections after the checkpoint.
                    _errors.ThrowIfFailed(includeRejections: false);
                    _engine.Advanced.ProcessTasks();
                    _errors.ThrowIfFailed();
                    T result = action();
                    // Native DOM actions can invoke JS callbacks without entering Evaluate.
                    // Complete their promise jobs before admitting the next session command.
                    _engine.Advanced.ProcessTasks();
                    _errors.ThrowIfFailed();
                    completion.TrySetResult(result);
                } catch (Exception error) {
                    // Preserve a listener's reported JS error when the provider rethrows an
                    // outer invocation failure with an empty or less useful message.
                    try { _errors.ThrowIfFailed(); }
                    catch (Exception tracked) { error = tracked; }
                    completion.TrySetException(error);
                }
            }
        }, TaskPriority.Normal);
        return completion.Task.WaitAsync(token);
    }

    public void Dispose() {
        _scripting.Dispose();
        _fetch?.Dispose();
        _loop?.CancelAll();
        _resources.Dispose();
        _context.Dispose();
    }
}
