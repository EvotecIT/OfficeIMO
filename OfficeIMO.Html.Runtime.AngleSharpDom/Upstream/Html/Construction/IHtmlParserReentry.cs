namespace AngleSharp.Html.Construction;

using System;

/// <summary>
/// Accepts source inserted while the HTML parser is paused in a script.
/// </summary>
internal interface IHtmlParserReentry
{
    /// <summary>
    /// Establishes an insertion point for a parser-blocking script.
    /// </summary>
    Boolean EnterScript();

    /// <summary>
    /// Restores the insertion point that preceded the current script.
    /// </summary>
    void ExitScript();

    /// <summary>
    /// Inserts and consumes the supplied source before returning to the caller.
    /// </summary>
    Boolean Write(String content);
}
