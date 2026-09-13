namespace AngleSharp.Io.Processors
{
    using AngleSharp.Dom;
    using AngleSharp.Html.Dom;
    using AngleSharp.Scripting;
    using AngleSharp.Text;
    using System;
    using System.Threading;
    using System.Threading.Tasks;

    sealed class ScriptRequestProcessor : IRequestProcessor
    {
        #region Fields

        private readonly IBrowsingContext _context;
        private readonly Document _document;
        private readonly HtmlScriptElement _script;
        private readonly IResourceLoader _loader;
        private IResponse? _response;
        private IScriptingService? _engine;
        private ScriptOptions? _options;
        private String? _inlineSource;

        #endregion

        #region ctor

        public ScriptRequestProcessor(IBrowsingContext context, HtmlScriptElement script)
        {
            _context = context;
            _document = script.Owner;
            _script = script;
            _loader = context.GetService<IResourceLoader>()!;
        }

        #endregion

        #region Properties

        public IDownload? Download
        {
            get;
            private set;
        }

        public IScriptingService? Engine => _engine ??= _context.GetScripting(ScriptLanguage);

        public String? AlternativeLanguage
        {
            get
            {
                var language = _script.GetOwnAttribute(AttributeNames.Language);
                return language != null ? "text/" + language : null;
            }
        }

        public String ScriptLanguage
        {
            get
            {
                var type = _script.Type ?? AlternativeLanguage;
                return type is { Length: > 0 } ? type : MimeTypeNames.DefaultJavaScript;
            }
        }

        #endregion

        #region Methods

        public async Task RunAsync(CancellationToken cancel)
        {
            var download = Download;

            if (download != null)
            {
                try
                {
                    _response = await download.Task.ConfigureAwait(false);
                }
                catch
                {
                    await _document.QueueTaskAsync(FireErrorEvent).ConfigureAwait(false);
                }
            }

            if (_response != null)
            {
                var response = _response;
                try
                {
                    var cancelled = await _document.QueueTaskAsync(FireBeforeScriptExecuteEvent).ConfigureAwait(false);
                    if (cancelled)
                    {
                        return;
                    }

                    var options = _options ?? CreateOptions();
                    var insert = _script.IsParserBlocking ? _document.Source.Index : -1;
                    var writeVersion = _document.ParserWriteVersion;
                    var previousScript = _document.CurrentScript;
                    var parserScript = _script.IsParserBlocking;
                    var ignoreDestructiveWrites = options.IsExternal || options.PreparedType.Isi("module");

                    var enteredParser = parserScript && _document.EnterParserScript();
                    if (ignoreDestructiveWrites) _document.EnterIgnoreDestructiveWrites();
                    _document.CurrentScript = options.PreparedType.Isi("module") ? null : _script;
                    try
                    {
                        await _engine!.EvaluateScriptAsync(response, options, cancel).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        /* We omit failed 3rd party services */
                        _context.TrackError(ex);
                    }
                    finally
                    {
                        _document.CurrentScript = previousScript;
                        if (ignoreDestructiveWrites) _document.ExitIgnoreDestructiveWrites();
                        if (enteredParser) _document.ExitParserScript();
                    }

                    // Async/deferred scripts must never rewind a parser that
                    // progressed while their evaluation was being scheduled.
                    if (insert >= 0 && writeVersion == _document.ParserWriteVersion) _document.Source.Index = insert;
                    await _document.QueueTaskAsync(FireAfterScriptExecuteEvent).ConfigureAwait(false);
                    await _document.QueueTaskAsync(FireLoadEvent).ConfigureAwait(false);
                }
                finally
                {
                    response.Dispose();
                    if (ReferenceEquals(_response, response)) _response = null;
                }
            }
        }

        public void Process(String content)
        {
            if (Engine != null)
            {
                _inlineSource = content;
                _options = CreateOptions();
                _response = VirtualResponse.Create(res => res.Content(content).Address(_script.BaseUri));
            }
        }

        public Boolean RunSynchronously()
        {
            if (_response is null || _inlineSource is null || Engine is not ISynchronousScriptingService synchronous)
            {
                return false;
            }

            var options = _options ?? CreateOptions();
            if (options.PreparedType.Isi("module"))
            {
                return false;
            }

            if (FireBeforeScriptExecuteEvent(CancellationToken.None))
            {
                _response.Dispose();
                _response = null;
                _inlineSource = null;
                return true;
            }

            var previousScript = _document.CurrentScript;
            var parserScript = _script.IsParserBlocking;
            var enteredParser = parserScript && _document.EnterParserScript();
            _document.CurrentScript = options.PreparedType.Isi("importmap") ? null : _script;
            try
            {
                var sourceUrl = RuntimeBaseUrl(options.Document);
                synchronous.EvaluateScript(options.Document, _inlineSource, options.PreparedType!, sourceUrl);
            }
            catch (Exception ex)
            {
                _context.TrackError(ex);
            }
            finally
            {
                _document.CurrentScript = previousScript;
                if (enteredParser) _document.ExitParserScript();
            }

            FireAfterScriptExecuteEvent(CancellationToken.None);
            FireLoadEvent(CancellationToken.None);
            _response.Dispose();
            _response = null;
            _inlineSource = null;
            return true;
        }

        public Task ProcessAsync(ResourceRequest request)
        {
            if (_loader != null && Engine != null)
            {
                _options = CreateOptions();
                _options.IsExternal = true;
                Download = _loader.FetchWithCorsAsync(new CorsRequest(request)
                {
                    Behavior = OriginBehavior.Taint,
                    Setting = _script.CrossOrigin.ToEnum(CorsSetting.None),
                    Integrity = _context.GetProvider<IIntegrityProvider>()
                });
                _options.PreparedSourceUrl = request.Target.Href;
                return Download.Task;
            }

            return Task.CompletedTask;
        }

        #endregion

        #region Helpers

        private ScriptOptions CreateOptions() => new(_document, _document.Loop!)
        {
            Element = _script,
            PreparedType = ScriptLanguage,
            Encoding = TextEncoding.Resolve(_script.CharacterSet)
        };

        private static String RuntimeBaseUrl(IDocument document) => document.BaseUri;

        private void FireLoadEvent(CancellationToken _) =>
            _script.FireSimpleEvent(EventNames.Load);

        private void FireErrorEvent(CancellationToken _) =>
            _script.FireSimpleEvent(EventNames.Error);

        private Boolean FireBeforeScriptExecuteEvent(CancellationToken _) =>
            _script.FireSimpleEvent(EventNames.BeforeScriptExecute, cancelable: true);

        private void FireAfterScriptExecuteEvent(CancellationToken _) =>
            _script.FireSimpleEvent(EventNames.AfterScriptExecute, bubble: true);

        #endregion
    }
}
