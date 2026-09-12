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

    private ScriptedDocumentSession(HtmlScriptRequest options) {
        _options = options;
        _errors = new RuntimeScriptErrors(options.MaxPendingPromiseRejections);
        _resources = new RuntimeResourceLoader(options);
        var configuration = Configuration.Default.WithCss().WithJs(new JsScriptingOptions { MaxCallStackDepth = 512 }).WithEventLoop()
            .With(new RuntimeResourceRequester(_resources, _errors))
            .WithDefaultLoader(new LoaderOptions { IsResourceLoadingEnabled = true, IsNavigationDisabled = true });
        var scripting = configuration.Services.OfType<JsScriptingService>().Single();
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
            var normalizeWindow = RuntimeWindowBindings.Install(_engine, document.DefaultView!);
            RuntimeEventBindings.Install(_engine, document.DefaultView!, _errors.Report, normalizeWindow);
            RuntimeUrlBindings.Install(_engine, document.DefaultView!);
            RuntimeObserverBindings.Install(_engine, document, _errors.Report);
            RuntimeStorageBindings.Install(_engine, options.MaxStorageCharacters);
            _fetch = new RuntimeFetchBindings(_engine, document, _loop, _resources, options, _errors);
        };
    }

    internal static async Task<ScriptedDocumentSession> OpenAsync(HtmlScriptRequest request, CancellationToken token) {
        var session = new ScriptedDocumentSession(request);
        try {
            session._document = await session._context.OpenAsync(source => source.Address(request.DocumentUrl.AbsoluteUri).Content(request.Html), token).WaitUntilAvailable(token);
            session._loop = session._context.GetService<IEventLoop>() ?? throw new HtmlScriptRuntimeException("The provider did not create an event loop.");
            foreach (string script in request.Scripts) await session.ExecuteAsync(script, token);
            await session.OnLoop(() => true, token);
            return session;
        } catch { session.Dispose(); throw; }
    }

    internal Task ExecuteAsync(string script, CancellationToken token) => OnLoop(() => { _document.ExecuteScript(script); return true; }, token);

    internal Task<string> EvaluateAsync(string expression, CancellationToken token) => OnLoop(() => {
        var value = _engine.Evaluate("JSON.stringify((" + expression + "\n))");
        if (!value.IsString()) throw new HtmlScriptRuntimeException("The expression did not produce a JSON value.");
        return value.AsString();
    }, token);

    internal async Task<HtmlRuntimeWireDocument?> WaitAsync(string expression, bool capture, CancellationToken token) {
        while (true) {
            token.ThrowIfCancellationRequested();
            var result = await OnLoop(() => {
                if (_document.ExecuteScript(expression) is not true) return (Ready: false, Document: (HtmlRuntimeWireDocument?)null);
                _errors.ThrowIfFailed();
                // Readiness and capture share a task so timers cannot mutate between them.
                var document = capture ? RuntimeDomCapture.Capture(_document, _options, token) : null;
                if (document != null) { document.DocumentUrl = new Uri(_document.Url); document.Resources = _resources.Capture().ToList(); }
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
                    _errors.ThrowIfFailed();
                    T result = action();
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
        _fetch?.Dispose();
        _loop?.CancelAll();
        _resources.Dispose();
        _context.Dispose();
    }
}
