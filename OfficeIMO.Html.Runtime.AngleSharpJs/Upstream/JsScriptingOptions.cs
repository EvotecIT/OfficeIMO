namespace AngleSharp.Js
{
    using System;

    /// <summary>
    /// Options tuning the JavaScript engine.
    /// </summary>
    public sealed class JsScriptingOptions
    {
        /// <summary>Configures host services before the interpreter is created for a window.</summary>
        public Action<AngleSharp.Dom.IWindow, Jint.Options> ConfigureEngine { get; set; }

        /// <summary>
        /// Gets or sets the JavaScript call stack depth that has to be supported
        /// before the engine gives up with a "Maximum call stack size exceeded"
        /// error. Defaults to 10000, which is roughly what browsers allow. Values
        /// of zero or less remove the limit - the engine is then bounded by the
        /// native stack alone, so a runaway recursion terminates the process with
        /// an uncatchable StackOverflowException.
        /// </summary>
        public Int32 MaxCallStackDepth { get; set; } = 10000;

        //  An engine is built per window, long after the options were handed over, so
        //  reading them then would let a later edit of the caller's object decide how
        //  the next document behaves. The service takes this copy instead.
        internal JsScriptingOptions Clone() => new JsScriptingOptions
        {
            MaxCallStackDepth = MaxCallStackDepth,
            ConfigureEngine = ConfigureEngine,
        };
    }
}
