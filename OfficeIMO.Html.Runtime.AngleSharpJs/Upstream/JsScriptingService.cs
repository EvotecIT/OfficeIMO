namespace AngleSharp.Scripting
{
    using AngleSharp.Browser;
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Js;
    using AngleSharp.Text;
    using Jint;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// The JavaScript engine.
    /// </summary>
    public class JsScriptingService : IScriptingService
    {
        #region Fields

        private readonly ConditionalWeakTable<IWindow, EngineInstance> _contexts;
        private readonly Dictionary<String, Object> _external;
        private readonly JsScriptingOptions _options;

        #endregion

        #region ctor

        /// <summary>
        /// Creates a new JavaScript engine using the default options.
        /// </summary>
        public JsScriptingService()
            : this(new JsScriptingOptions())
        {
        }

        /// <summary>
        /// Creates a new JavaScript engine using the given options. The options
        /// are copied, so that editing them afterwards leaves this engine alone.
        /// </summary>
        /// <param name="options">The options tuning the engine.</param>
        public JsScriptingService(JsScriptingOptions options)
        {
            _contexts = new ConditionalWeakTable<IWindow, EngineInstance>();
            _external = new Dictionary<String, Object>();
            _options = (options ?? throw new ArgumentNullException(nameof(options))).Clone();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the external assignments.
        /// </summary>
        public IDictionary<String, Object> External => _external;

        /// <summary>
        /// Gets the engine's mime-type.
        /// </summary>
        public String Type => MimeTypeNames.DefaultJavaScript;

        #endregion

        #region Methods

        Boolean IScriptingService.SupportsType(String mimeType) =>
            MimeTypeNames.IsJavaScript(mimeType) ||
            mimeType.Isi("module") ||
            mimeType.Isi("importmap");

        /// <summary>
        /// Gets the associated Jint engine or creates it.
        /// </summary>
        /// <param name="document">The current document.</param>
        /// <returns>The engine object.</returns>
        public Engine GetOrCreateJint(IDocument document) =>
            GetOrCreateInstance(document).Jint;

        /// <summary>
        /// Evaluates the response asynchronously.
        /// </summary>
        /// <param name="response">The response to parse.</param>
        /// <param name="options">The options to consider.</param>
        /// <param name="cancel">The cancellation token to transport.</param>
        public async Task EvaluateScriptAsync(IResponse response, ScriptOptions options, CancellationToken cancel)
        {
            var encoding = options.Encoding ?? Encoding.UTF8;

            using (var reader = new StreamReader(response.Content, encoding, true))
            {
                var content = await reader.ReadToEndAsync().ConfigureAwait(false);
                await options.EventLoop.EnqueueAsync(_ =>
                    EvaluateScript(options.Document, content, options.Element?.Type, options.Element?.Source), TaskPriority.Critical).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Evaluates the given script source in the engine of the document.
        /// </summary>
        /// <param name="document">The context of the evaluation.</param>
        /// <param name="source">The source of the script.</param>
        /// <param name="type">The type of the script.</param>
        /// <param name="sourceUrl">The URL of the script.</param>
        /// <returns>The result of the evaluation.</returns>
        public Object EvaluateScript(IDocument document, String source, String type, String sourceUrl)
        {
            document = document ?? throw new ArgumentNullException(nameof(document));
            return GetOrCreateInstance(document).RunScript(source, type, sourceUrl).FromJsValue();
        }

        #endregion

        #region Helpers

        internal EngineInstance GetOrCreateInstance(IDocument document)
        {
            var objectContext = document.DefaultView;

            if (!_contexts.TryGetValue(objectContext, out var instance))
            {
                var libs = GetAssemblies(document.Context).ToArray();
                instance = new EngineInstance(objectContext, _external, libs, _options);
                _contexts.Add(objectContext, instance);
            }

            return instance;
        }

        private static IEnumerable<Assembly> GetAssemblies(IBrowsingContext context) => context
            .GetServices<Object>()
            .Select(m => m.GetType().GetAssembly())
            .Distinct()
            .Where(m => m.FullName.StartsWith("AngleSharp"));

        #endregion
    }
}
