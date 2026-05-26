using System;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level org chart authoring extensions for <see cref="VisioDocument"/>.
    /// </summary>
    public static class VisioOrgChartDiagramDocumentExtensions {
        /// <summary>
        /// Adds a semantic org chart page and returns the document for chaining.
        /// </summary>
        public static VisioDocument OrgChartDiagram(this VisioDocument document, string pageName, Action<VisioOrgChartDiagramBuilder> configure) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            VisioOrgChartDiagramBuilder builder = new VisioOrgChartDiagramBuilder(document, pageName);
            configure?.Invoke(builder);
            builder.Build();
            return document;
        }
    }
}
