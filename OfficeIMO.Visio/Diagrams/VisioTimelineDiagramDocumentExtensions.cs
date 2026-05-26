using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level timeline authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioTimelineDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic timeline page and returns the document for chaining.
        /// </summary>
        public static VisioDocument TimelineDiagram(this VisioDocument document, string pageName, Action<VisioTimelineDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioTimelineDiagramBuilder builder = new VisioTimelineDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
