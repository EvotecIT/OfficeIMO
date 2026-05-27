using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level generic graph diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioGraphDiagramDocumentExtensions {
        /// <summary>
        /// Adds an automatically laid out generic graph page and returns the document for chaining.
        /// </summary>
        public static VisioDocument GraphDiagram(this VisioDocument document, string pageName, Action<VisioGraphDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioGraphDiagramBuilder builder = new VisioGraphDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
