using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level diagram authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioFlowchartDocumentExtensions {
        /// <summary>
        /// Adds a semantic flowchart page and returns the document for chaining.
        /// </summary>
        public static VisioDocument Flowchart(this VisioDocument document, string pageName, Action<VisioFlowchartBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioFlowchartBuilder builder = new VisioFlowchartBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
