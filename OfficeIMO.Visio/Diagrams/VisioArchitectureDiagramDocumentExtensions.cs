using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level architecture diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioArchitectureDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic architecture diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument ArchitectureDiagram(this VisioDocument document, string pageName, Action<VisioArchitectureDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioArchitectureDiagramBuilder builder = new VisioArchitectureDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
