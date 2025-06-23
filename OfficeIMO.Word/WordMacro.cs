using System;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a single macro module within a document.
    /// </summary>
    /// <remarks>
    /// Instances are returned by <see cref="WordDocument.Macros"/> and can be
    /// removed individually using <see cref="Remove"/>.
    /// </remarks>
    public class WordMacro {
        private readonly WordDocument _document;

        /// <summary>
        /// Gets the macro module name.
        /// </summary>
        public string Name { get; }

        internal WordMacro(WordDocument document, string name) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        /// <summary>
        /// Removes this macro module from the document.
        /// </summary>
        public void Remove() {
            _document.RemoveMacro(Name);
        }
    }
}
