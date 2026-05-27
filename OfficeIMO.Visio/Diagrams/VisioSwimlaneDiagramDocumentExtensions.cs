using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level swimlane/process map authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioSwimlaneDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic swimlane diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument SwimlaneDiagram(this VisioDocument document, string pageName, Action<VisioSwimlaneDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioSwimlaneDiagramBuilder builder = new VisioSwimlaneDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
