using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level dependency diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioDependencyDiagramDocumentExtensions {
        /// <summary>
        /// Adds an automatically layered dependency diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument DependencyDiagram(this VisioDocument document, string pageName, Action<VisioDependencyDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioDependencyDiagramBuilder builder = new VisioDependencyDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
