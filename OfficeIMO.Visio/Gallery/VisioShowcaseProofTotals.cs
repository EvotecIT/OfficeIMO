using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Aggregates structural proof metrics across every generated showcase diagram.
    /// </summary>
    public sealed class VisioShowcaseProofTotals {
        internal VisioShowcaseProofTotals(
            int totalShapeCount,
            int connectorCount,
            int stencilUsageCount,
            int totalConnectionPointCount,
            int connectionPointShapeCount,
            int stencilFamilyCount,
            int stencilBackedShapeCount,
            int basicGeometryShapeCount,
            int masterBackedShapeCount,
            int packageBackedShapeCount,
            int generatedMasterBackedShapeCount,
            int semanticOnlyShapeCount,
            IReadOnlyList<string> stencilCatalogs,
            IReadOnlyList<string> shapeDataKeys,
            IReadOnlyList<string> connectorShapeDataKeys,
            IReadOnlyList<string> semanticKinds) {
            TotalShapeCount = totalShapeCount;
            ConnectorCount = connectorCount;
            StencilUsageCount = stencilUsageCount;
            TotalConnectionPointCount = totalConnectionPointCount;
            ConnectionPointShapeCount = connectionPointShapeCount;
            StencilFamilyCount = stencilFamilyCount;
            StencilBackedShapeCount = stencilBackedShapeCount;
            BasicGeometryShapeCount = basicGeometryShapeCount;
            MasterBackedShapeCount = masterBackedShapeCount;
            PackageBackedShapeCount = packageBackedShapeCount;
            GeneratedMasterBackedShapeCount = generatedMasterBackedShapeCount;
            SemanticOnlyShapeCount = semanticOnlyShapeCount;
            StencilCatalogs = stencilCatalogs;
            ShapeDataKeys = shapeDataKeys;
            ConnectorShapeDataKeys = connectorShapeDataKeys;
            SemanticKinds = semanticKinds;
        }

        /// <summary>Total shape count across diagram stencil profiles.</summary>
        public int TotalShapeCount { get; }

        /// <summary>Total connector count across diagram stencil profiles.</summary>
        public int ConnectorCount { get; }

        /// <summary>Total grouped stencil/profile usage count across diagram stencil profiles.</summary>
        public int StencilUsageCount { get; }

        /// <summary>Total connection point count across diagram stencil profiles.</summary>
        public int TotalConnectionPointCount { get; }

        /// <summary>Total number of shapes with connection points across diagram stencil profiles.</summary>
        public int ConnectionPointShapeCount { get; }

        /// <summary>Total stencil family count across diagram stencil profiles.</summary>
        public int StencilFamilyCount { get; }

        /// <summary>Total stencil-backed shape count across diagram stencil profiles.</summary>
        public int StencilBackedShapeCount { get; }

        /// <summary>Total basic geometry shape count across diagram stencil profiles.</summary>
        public int BasicGeometryShapeCount { get; }

        /// <summary>Total master-backed shape count across diagram stencil profiles.</summary>
        public int MasterBackedShapeCount { get; }

        /// <summary>Total package-backed shape count across diagram stencil profiles.</summary>
        public int PackageBackedShapeCount { get; }

        /// <summary>Total generated-master-backed shape count across diagram stencil profiles.</summary>
        public int GeneratedMasterBackedShapeCount { get; }

        /// <summary>Total semantic-only shape count across diagram stencil profiles.</summary>
        public int SemanticOnlyShapeCount { get; }

        /// <summary>Distinct stencil catalogs represented across all diagram stencil profiles.</summary>
        public IReadOnlyList<string> StencilCatalogs { get; }

        /// <summary>Number of distinct stencil catalogs represented across all diagram stencil profiles.</summary>
        public int StencilCatalogCount => StencilCatalogs.Count;

        /// <summary>Distinct Shape Data keys represented across all diagram stencil profiles.</summary>
        public IReadOnlyList<string> ShapeDataKeys { get; }

        /// <summary>Number of distinct Shape Data keys represented across all diagram stencil profiles.</summary>
        public int ShapeDataKeyCount => ShapeDataKeys.Count;

        /// <summary>Distinct connector Shape Data keys represented across all diagram stencil profiles.</summary>
        public IReadOnlyList<string> ConnectorShapeDataKeys { get; }

        /// <summary>Number of distinct connector Shape Data keys represented across all diagram stencil profiles.</summary>
        public int ConnectorShapeDataKeyCount => ConnectorShapeDataKeys.Count;

        /// <summary>Distinct semantic kind values represented across all diagram stencil profiles.</summary>
        public IReadOnlyList<string> SemanticKinds { get; }

        /// <summary>Number of distinct semantic kind values represented across all diagram stencil profiles.</summary>
        public int SemanticKindCount => SemanticKinds.Count;
    }
}
