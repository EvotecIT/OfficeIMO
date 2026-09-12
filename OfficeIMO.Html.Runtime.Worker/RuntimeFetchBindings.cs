using System.Text;
using System.Text.Json;
using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;
using Jint.Runtime;

namespace OfficeIMO.Html.Runtime.Worker;

// Async transport never touches the interpreter. Only event-loop tasks settle JS promises.
internal sealed class RuntimeFetchBindings : IDisposable {
    private readonly Engine _engine;
    private readonly IDocument _document;
    private readonly IEventLoop _loop;
    private readonly RuntimeResourceLoader _loader;
    private readonly HtmlScriptRequest _options;
    private readonly RuntimeScriptErrors _errors;
    private readonly Dictionary<int, CancellationTokenSource> _pending = new();
    private readonly object _sync = new();
    private int _nextId;
    private bool _disposed;

    internal RuntimeFetchBindings(Engine engine, IDocument document, IEventLoop loop, RuntimeResourceLoader loader, HtmlScriptRequest options, RuntimeScriptErrors errors) {
        _engine = engine; _document = document; _loop = loop; _loader = loader; _options = options; _errors = errors;
        using var stream = typeof(RuntimeFetchBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeFetchBootstrap.js")!;
        using var reader = new StreamReader(stream);
        JsValue factory = engine.Evaluate(reader.ReadToEnd());
        var start = new ClrFunction(engine, "start", (_, args) => AsJavaScript(() => Start(args[0].AsString(), args[1])));
        var cancel = JsValue.FromObject(engine, (Action<int>)Cancel);
        var encode = new ClrFunction(engine, "encode", (_, args) => AsJavaScript(() => Encode(args[0].AsString())));
        var decode = JsValue.FromObject(engine, (Func<string, string>)(encoded => {
            string value = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
            return value.StartsWith('\uFEFF') ? value[1..] : value;
        }));
        var report = JsValue.FromObject(engine, (Action<string>)errors.Report);
        var exports = engine.Invoke(factory, new[] { start, cancel, encode, decode, report, JsValue.FromObject(engine, options.ResourcePolicy.MaxRequestBytes) }).AsObject();
        var window = JsValue.FromObject(engine, document.DefaultView).AsObject();
        foreach (var property in exports.GetOwnProperties()) {
            var descriptor = new PropertyDescriptor(property.Value.Value, true, false, true);
            engine.Global.FastSetProperty(property.Key, descriptor);
            window.FastSetProperty(property.Key, descriptor);
        }
    }

    private JsValue AsJavaScript(Func<JsValue> action) {
        try { return action(); }
        catch (Exception error) { throw new JavaScriptException(_engine.Intrinsics.TypeError, error.Message); }
    }

    private string Encode(string value) {
        if (Encoding.UTF8.GetByteCount(value) > _options.ResourcePolicy.MaxRequestBytes) throw new HtmlScriptRuntimeException("Fetch request body exceeds its byte budget.");
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
    }

    private JsValue Start(string json, JsValue callback) {
        if (json.Length > _options.MaxInputCharacters) throw new HtmlScriptRuntimeException("Fetch request exceeds the input character budget.");
        var request = JsonSerializer.Deserialize<RuntimeFetchRequest>(json) ?? throw new HtmlScriptRuntimeException("Invalid fetch request.");
        // Read live base URI only on the interpreter thread; the document's origin is fixed by the loader.
        var url = new Uri(new Uri(_document.BaseUri), request.Url);
        CancellationTokenSource cancellation;
        int id;
        lock (_sync) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_nextId >= _options.ResourcePolicy.MaxRequests) throw new HtmlScriptRuntimeException("Fetch request budget exceeded.");
            id = ++_nextId;
            cancellation = new CancellationTokenSource();
            _pending.Add(id, cancellation);
        }
        _ = CompleteAsync(id, url, request, callback, cancellation);
        return id;
    }

    private async Task CompleteAsync(int id, Uri url, RuntimeFetchRequest request, JsValue callback, CancellationTokenSource cancellation) {
        string json;
        try {
            var response = await _loader.FetchAsync(url, request, cancellation.Token).ConfigureAwait(false);
            bool cors = HtmlRuntimeResourcePolicy.Origin(response.FinalUrl) != HtmlRuntimeResourcePolicy.Origin(_options.DocumentUrl);
            bool hasBody = request.Method != "HEAD" && response.StatusCode is not (204 or 205 or 304);
            json = JsonSerializer.Serialize(new {
                status = response.StatusCode, statusText = response.StatusText, url = response.FinalUrl.AbsoluteUri,
                redirected = response.RedirectCount > 0, type = cors ? "cors" : "basic", headers = RuntimeFetchCors.Expose(response, cors),
                body = hasBody ? Convert.ToBase64String(response.Buffer) : "", hasBody
            });
        } catch (Exception error) {
            json = JsonSerializer.Serialize(new { error = error is OperationCanceledException ? "AbortError" : "TypeError", message = error.Message });
        }
        lock (_sync) {
            if (_disposed) { _pending.Remove(id); cancellation.Dispose(); return; }
            _loop.Enqueue(_ => {
                lock (_engine) {
                    lock (_sync) {
                        if (_disposed) return;
                        _pending.Remove(id);
                        cancellation.Dispose();
                    }
                    try {
                        _engine.Invoke(callback, new JsValue[] { json });
                    } catch (Exception error) { _errors.Report(error.Message); }
                }
            }, TaskPriority.Normal);
        }
    }

    private void Cancel(int id) { lock (_sync) { if (_pending.TryGetValue(id, out var token)) token.Cancel(); } }
    public void Dispose() {
        lock (_sync) {
            _disposed = true;
            foreach (var token in _pending.Values) token.Cancel();
            // Completion tasks dispose sources after observing cancellation.
        }
    }
}
