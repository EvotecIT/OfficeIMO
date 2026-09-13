namespace AngleSharp.Scripting
{
    using AngleSharp.Dom;
    using System;

    /// <summary>
    /// Provides the synchronous classic-script entry point required while an
    /// HTML parser is reentered by document.write or a DOM-inserted inline script.
    /// </summary>
    public interface ISynchronousScriptingService
    {
        /// <summary>
        /// Evaluates a prepared, non-module script on the current script stack.
        /// </summary>
        Object EvaluateScript(IDocument document, String source, String type, String sourceUrl);
    }
}
