namespace AngleSharp.Js
{
    using AngleSharp.Browser;
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Scripting;
    using AngleSharp.Text;
    using System;
    using System.IO;
    using System.Threading;
    using System.Threading.Tasks;

    /// <summary>
    /// Represents a handler for javascript: URLs.
    /// </summary>
    public class JsNavigationHandler : INavigationHandler
    {
        private static readonly String UrlSchema = "javascript";
        private readonly JsScriptingService _service;

        /// <summary>
        /// Creates a new navigation handler for javascript: URLs.
        /// </summary>
        /// <param name="service">The underlying scripting service.</param>
        public JsNavigationHandler(JsScriptingService service)
        {
            _service = service;
        }

        /// <inheritdoc />
        public async Task<IDocument> NavigateAsync(DocumentRequest request, CancellationToken token)
        {
            var document = request.Source?.Owner;

            if (document != null)
            {
                var loop = document.Context.GetService<IEventLoop>();
                var options = new ScriptOptions(document, loop);
                var script = request.Target.Href.Substring(UrlSchema.Length + 1);
                var content = TextEncoding.Utf8.GetBytes(script);
                var response = new DefaultResponse
                {
                    Address = new Url(document.Url),
                    Content = new MemoryStream(content),
                };
                await _service.EvaluateScriptAsync(response, options, token).ConfigureAwait(false);
            }

            return document;
        }

        /// <inheritdoc />
        public Boolean SupportsProtocol(String protocol) => protocol.Is(UrlSchema);
    }
}
