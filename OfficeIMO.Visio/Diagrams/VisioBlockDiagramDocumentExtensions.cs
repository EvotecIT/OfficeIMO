using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level block/system diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioBlockDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic block diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument BlockDiagram(this VisioDocument document, string pageName, Action<VisioBlockDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioBlockDiagramBuilder builder = new VisioBlockDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
