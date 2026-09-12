namespace AngleSharp.Js.Cache
{
    using Acornima.Ast;
    using Jint;
    using Jint.Runtime;
    using System;
    using System.Collections.Concurrent;

    /// <summary>
    /// Caches the parsed and analyzed form of scripts. A prepared script is
    /// documented by Jint as reusable and thread-safe, so the same one can serve
    /// every document that runs the same source - a page loading a library ends up
    /// parsing it once instead of once per document.
    /// </summary>
    static class ScriptCache
    {
        //  There is no bound on how many distinct scripts a process may see, so the
        //  cache does not grow without end. The scripts worth keeping are the ones
        //  seen first: libraries are referenced early and by many documents.
        private const Int32 Capacity = 32;

        private static readonly ConcurrentDictionary<String, Prepared<Script>> _scripts =
            new ConcurrentDictionary<String, Prepared<Script>>(StringComparer.Ordinal);

        /// <summary>
        /// Gets the prepared form of the given source. An invalid result means the
        /// source could not be prepared and should be handed to the engine as text,
        /// so that the engine reports the syntax error itself.
        /// </summary>
        public static Prepared<Script> GetOrCreate(String source)
        {
            if (_scripts.TryGetValue(source, out var prepared))
            {
                return prepared;
            }

            try
            {
                prepared = Engine.PrepareScript(source);
            }
            catch (ScriptPreparationException)
            {
                return default;
            }

            if (_scripts.Count < Capacity)
            {
                _scripts.TryAdd(source, prepared);
            }

            return prepared;
        }
    }
}
