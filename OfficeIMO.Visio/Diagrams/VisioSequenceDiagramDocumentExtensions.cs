using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level sequence diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioSequenceDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic sequence diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument SequenceDiagram(this VisioDocument document, string pageName, Action<VisioSequenceDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioSequenceDiagramBuilder builder = new VisioSequenceDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
