using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level network diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioNetworkDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic network diagram page and returns the document for chaining.
        /// </summary>
        public static VisioDocument NetworkDiagram(this VisioDocument document, string pageName, Action<VisioNetworkDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioNetworkDiagramBuilder builder = new VisioNetworkDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }

        /// <summary>
        /// Adds a semantic network topology page and returns the document for chaining.
        /// </summary>
        public static VisioDocument NetworkTopologyDiagram(this VisioDocument document, string pageName, Action<VisioNetworkTopologyDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioNetworkTopologyDiagramBuilder builder = new VisioNetworkTopologyDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
