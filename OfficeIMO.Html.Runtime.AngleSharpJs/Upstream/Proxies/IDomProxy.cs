namespace AngleSharp.Js
{
    using System;

    /// <summary>
    /// A JS object standing for one CLR DOM object.
    /// </summary>
    /// <remarks>
    /// There are two of them and they cannot share a base class: an indexed collection derives
    /// from Jint's array-like base so that the engine knows how to read it, everything else from
    /// the ordinary object base. What every caller outside the proxies needs is the same either
    /// way - the CLR object behind the proxy - so that is what they agree on.
    /// </remarks>
    interface IDomProxy
    {
        /// <summary>
        /// Gets the DOM object this proxy stands for.
        /// </summary>
        Object Value { get; }

        /// <summary>
        /// Gets the engine the proxy belongs to. A prototype member is shared by every engine
        /// and reaches its own through whichever object it was invoked on.
        /// </summary>
        EngineInstance Instance { get; }
    }
}
