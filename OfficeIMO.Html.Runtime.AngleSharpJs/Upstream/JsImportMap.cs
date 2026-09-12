namespace AngleSharp.Js
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// https://html.spec.whatwg.org/multipage/webappapis.html#import-map
    /// </summary>
    sealed class JsImportMap
    {
        public JsImportMap()
        {
            Imports = new Dictionary<string, Uri>();
            Scopes = new Dictionary<string, Dictionary<string, Uri>>();
        }

        /// <summary>
        /// Provides the mappings between module specifier text that might appear in an import statement or import() operator,
        /// and the text that will replace it when the specifier is resolved.
        /// </summary>
        public Dictionary<string, Uri> Imports { get; set; }

        /// <summary>
        /// Mappings that are only used if the script importing the module contains a particular URL path.
        /// If the URL of the loading script matches the supplied path, the mapping associated with the scope will be used.
        /// </summary>
        public Dictionary<string, Dictionary<string, Uri>> Scopes { get; set; }
    }
}
