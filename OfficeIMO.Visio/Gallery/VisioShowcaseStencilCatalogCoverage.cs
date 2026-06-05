using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Describes which generated showcase diagrams exercise a stencil catalog.
    /// </summary>
    public sealed class VisioShowcaseStencilCatalogCoverage {
        internal VisioShowcaseStencilCatalogCoverage(string catalog, IReadOnlyList<string> diagramNames) {
            Catalog = catalog;
            DiagramNames = diagramNames;
        }

        /// <summary>Stencil catalog name represented by one or more showcase diagrams.</summary>
        public string Catalog { get; }

        /// <summary>Generated showcase diagram names that use the stencil catalog.</summary>
        public IReadOnlyList<string> DiagramNames { get; }

        /// <summary>Number of generated showcase diagrams that use the stencil catalog.</summary>
        public int DiagramCount => DiagramNames.Count;
    }
}
