namespace AngleSharp.Js
{
    using AngleSharp.Dom;
    using AngleSharp.Io;
    using AngleSharp.Scripting;
    using System;

    /// <summary>
    /// Useful extensions for the DOM.
    /// </summary>
    public static class JsApiExtensions
    {
        /// <summary>
        /// Executes the given script code in the context of the document.
        /// </summary>
        /// <param name="document">The document as context.</param>
        /// <param name="scriptCode">The script to run.</param>
        /// <param name="scriptType">The type of the script to run (defaults to "text/javascript").</param>
        /// <param name="sourceUrl">The URL of the script.</param>
        /// <returns>The result of running the script, if any.</returns>
        public static Object ExecuteScript(this IDocument document, String scriptCode, String scriptType = null, String sourceUrl = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var service = document?.Context.GetService<JsScriptingService>();
            return service?.EvaluateScript(document, scriptCode, scriptType ?? MimeTypeNames.DefaultJavaScript, sourceUrl);
        }
    }
}
