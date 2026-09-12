namespace AngleSharp.Js.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Browser;
    using AngleSharp.Dom;
    using AngleSharp.Dom.Events;
    using AngleSharp.Io;
    using AngleSharp.Text;
    using AngleSharp.Scripting;
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Represents a dedicated worker as seen from a window.
    /// </summary>
    [DomName("Worker")]
    [DomExposed("Window")]
    public sealed class Worker : EventTarget
    {
        #region Fields

        private const String MessageEventName = "message";
        private const String WorkerIncomingMessageName = "__anglesharpIncomingMessage";

        private readonly IWindow _window;
        private readonly IBrowsingContext _workerContext;
        private readonly JsScriptingService _scripting;
        private readonly IEventLoop _windowLoop;
        private readonly IEventLoop _workerLoop;
        private readonly Queue<Object> _pendingMessages;

        private IDocument _workerDocument;
        private Boolean _initialized;
        private Exception _startupError;
        private Boolean _terminated;

        #endregion

        #region ctor

        /// <summary>
        /// Creates a new worker and starts the worker script in a separate context.
        /// </summary>
        [DomConstructor]
        public Worker(IWindow window, String source)
        {
            if (window == null)
            {
                throw new ArgumentNullException(nameof(window));
            }

            if (String.IsNullOrEmpty(source))
            {
                throw new ArgumentException("The worker source URL cannot be empty.", nameof(source));
            }

            _window = window;
            _windowLoop = _window.Document.Context.GetService<IEventLoop>();

            var parentContext = _window.Document.Context;
            _scripting = parentContext.GetService<JsScriptingService>() ?? throw new DomException(DomError.NotSupported);
            _workerLoop = new JsEventLoop();
            var workerConfig = Configuration.Default
                .With(_scripting)
                .WithOnly(_workerLoop);
            _workerContext = BrowsingContext.New(workerConfig);

            _pendingMessages = new Queue<Object>();

            EnsureLoaderAvailable();

            var workerUrl = ResolveUrl(source);
            Enqueue(_workerLoop, TaskPriority.Critical, () =>
            {
                try
                {
                    InitializeWorker(workerUrl);
                }
                catch (Exception ex)
                {
                    _startupError = ex;
                }
            });
        }

        #endregion

        #region Properties

        internal Exception StartupError => _startupError;

        internal Boolean IsInitialized => _initialized;

        internal Object EvaluateInWorker(String source)
        {
            if (_workerDocument == null)
            {
                return null;
            }

            return _scripting.EvaluateScript(_workerDocument, source, MimeTypeNames.DefaultJavaScript, _workerDocument.Url);
        }

        #endregion

        #region Events

        /// <summary>
        /// Adds or removes the handler for the message event.
        /// </summary>
        [DomName("onmessage")]
        public event DomEventHandler Message
        {
            add { AddEventListener(MessageEventName, value, false); }
            remove { RemoveEventListener(MessageEventName, value, false); }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Posts a message to the running worker script.
        /// </summary>
        [DomName("postMessage")]
        public void PostMessage(Object message)
        {
            if (_terminated)
            {
                return;
            }

            if (_startupError != null)
            {
                throw _startupError;
            }

            if (!_initialized)
            {
                lock (_pendingMessages)
                {
                    _pendingMessages.Enqueue(message);
                }

                return;
            }

            Enqueue(_workerLoop, TaskPriority.Normal, () => DeliverToWorker(message));
        }

        /// <summary>
        /// Terminates the running worker.
        /// </summary>
        [DomName("terminate")]
        public void Terminate()
        {
            _terminated = true;

            if (_workerLoop is IDisposable disposable)
            {
                disposable.Dispose();
            }
        }

        #endregion

        #region Helpers

        internal void PostMessageToOwner(Object message)
        {
            if (_terminated)
            {
                return;
            }

            var origin = _workerDocument?.Origin ?? String.Empty;
            var ev = new MessageEvent(MessageEventName, false, false, message, origin, String.Empty, null, null);
            Enqueue(_windowLoop, TaskPriority.Normal, () => Dispatch(ev));
        }

        private void InitializeWorker(Url sourceUrl)
        {
            _workerDocument = _workerContext.OpenAsync(request => request.Content("<!doctype html>")).GetAwaiter().GetResult();
            WindowExtensions.RegisterWorkerWindow(_workerDocument.DefaultView, this);
            var source = FetchWorkerScript(sourceUrl);
            _scripting.EvaluateScript(_workerDocument, WorkerBootstrap, MimeTypeNames.DefaultJavaScript, sourceUrl.Href);
            _scripting.EvaluateScript(_workerDocument, source, MimeTypeNames.DefaultJavaScript, sourceUrl.Href);

            _initialized = true;

            lock (_pendingMessages)
            {
                while (_pendingMessages.Count > 0)
                {
                    DeliverToWorker(_pendingMessages.Dequeue());
                }
            }
        }

        private void DeliverToWorker(Object message)
        {
            var engine = _scripting.GetOrCreateJint(_workerDocument);
            engine.SetValue(WorkerIncomingMessageName, message);
            _scripting.EvaluateScript(_workerDocument, WorkerDispatchScript, MimeTypeNames.DefaultJavaScript, _workerDocument.Url);
        }

        private String FetchWorkerScript(Url source)
        {
            var loader = _window.Document.Context.GetService<IDocumentLoader>();

            if (loader == null)
            {
                throw new DomException(DomError.NotSupported);
            }

            var request = new DocumentRequest(source)
            {
                Referer = _window.Document.DocumentUri,
            };

            using (var response = loader.FetchAsync(request).Task.GetAwaiter().GetResult())
            {
                if (response?.Content == null)
                {
                    throw new DomException(DomError.NotSupported);
                }

                using (var reader = new StreamReader(response.Content))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private void EnsureLoaderAvailable()
        {
            var loader = _window.Document.Context.GetService<IDocumentLoader>();

            if (loader == null)
            {
                throw new DomException(DomError.NotSupported);
            }
        }

        private Url ResolveUrl(String source)
        {
            if (source.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                source.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return Url.Create(source);
            }

            var documentUrl = _window.Document.Url;

            if (String.IsNullOrEmpty(documentUrl) || documentUrl.Is("about:blank"))
            {
                documentUrl = "https://example.com/";
            }

            return new Url(new Url(documentUrl), source);
        }

        private static void Enqueue(IEventLoop loop, TaskPriority priority, Action action)
        {
            if (loop != null)
            {
                loop.Enqueue(() => action(), priority);
            }
            else
            {
                action();
            }
        }

        private const String WorkerBootstrap = @"(function () {
if (typeof self === 'undefined') {
    this.self = this;
}
})();";

        private const String WorkerDispatchScript = @"(function () {
if (self && typeof self.onmessage === 'function') {
    self.onmessage({ data: __anglesharpIncomingMessage });
}
    })();";

        #endregion

    }
}