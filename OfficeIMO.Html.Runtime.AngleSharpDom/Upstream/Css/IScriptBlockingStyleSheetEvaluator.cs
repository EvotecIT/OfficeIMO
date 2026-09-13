namespace AngleSharp.Css
{
    using System;

    /// <summary>
    /// Evaluates stylesheet media conditions for parser script blocking.
    /// </summary>
    public interface IScriptBlockingStyleSheetEvaluator
    {
        /// <summary>
        /// Returns whether the media query list matches the active environment.
        /// </summary>
        Boolean Matches(String mediaText);
    }
}
