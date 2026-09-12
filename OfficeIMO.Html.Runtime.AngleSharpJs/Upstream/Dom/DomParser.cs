namespace AngleSharp.Js.Dom
{
    using AngleSharp.Attributes;
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Text;
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// A way for JavaScript applications to access the parser.
    /// See: https://w3c.github.io/DOM-Parsing/#the-domparser-interface
    /// </summary>
    [DomName("DOMParser")]
    [DomExposed("Window")]
    public sealed class DomParser
    {
        private readonly IWindow _window;

        /// <summary>
        /// Creates a new DOMParser instance.
        /// </summary>
        /// <param name="window">The window to host the parser.</param>
        [DomConstructor]
        public DomParser(IWindow window)
        {
            _window = window;
        }

        /// <summary>
        /// Parses the given string for the given type. Throws a not supported
        /// exception if the type is not supported.
        /// </summary>
        /// <param name="str">The content to parse.</param>
        /// <param name="type">The type of the target to parse to.</param>
        /// <returns>The document of the given type.</returns>
        [DomName("parseFromString")]
        public IDocument Parse(String str, String type)
        {
            var ctx = _window?.Document.Context;
            var factory = ctx?.GetService<IDocumentFactory>() ?? throw new DomException(DomError.NotSupported);

            using (var content = new MemoryStream(TextEncoding.Utf8.GetBytes(str)))
            {
                var response = new DefaultResponse
                {
                    Address = new Url(_window.Document.Url),
                    Content = content,
                    Headers = new Dictionary<String, String>
                    {
                        { HeaderNames.ContentType, type },
                    },
                };

                try
                {
                    var task = factory.CreateAsync(ctx, new CreateDocumentOptions(response), default);
                    var document = task.Result;
                    return document;
                }
                catch (Exception ex)
                {
                    var message = ex.InnerException?.Message ?? ex.Message ?? "Something went wrong.";
                    var sourceText = String.Empty;
                    var error = $"<parsererror xmlns=\"http://www.mozilla.org/newlayout/xml/parsererror.xml\">{message}<sourcetext>{sourceText}</sourcetext></parsererror>";
                    return Parse(error, MimeTypeNames.Xml);
                }
            }
        }
    }
}
