using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Summarizes structural proof details that are useful when reviewing a generated showcase diagram.
    /// </summary>
    public sealed class VisioShowcaseProofSummary {
        internal static readonly VisioShowcaseProofSummary Empty = new(
            Array.Empty<string>(),
            Array.Empty<string>(),
            Array.Empty<string>(),
            Array.Empty<string>(),
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0);

        internal VisioShowcaseProofSummary(
            IReadOnlyList<string> stencilCatalogs,
            IReadOnlyList<string> shapeDataKeys,
            IReadOnlyList<string> connectorShapeDataKeys,
            IReadOnlyList<string> semanticKinds,
            int stencilUsageCount,
            int totalShapeCount,
            int connectorCount,
            int totalConnectionPointCount,
            int connectionPointShapeCount,
            int stencilFamilyCount,
            int stencilBackedShapeCount,
            int basicGeometryShapeCount,
            int masterBackedShapeCount,
            int packageBackedShapeCount,
            int generatedMasterBackedShapeCount,
            int semanticOnlyShapeCount) {
            StencilCatalogs = stencilCatalogs;
            ShapeDataKeys = shapeDataKeys;
            ConnectorShapeDataKeys = connectorShapeDataKeys;
            SemanticKinds = semanticKinds;
            StencilUsageCount = stencilUsageCount;
            TotalShapeCount = totalShapeCount;
            ConnectorCount = connectorCount;
            TotalConnectionPointCount = totalConnectionPointCount;
            ConnectionPointShapeCount = connectionPointShapeCount;
            StencilFamilyCount = stencilFamilyCount;
            StencilBackedShapeCount = stencilBackedShapeCount;
            BasicGeometryShapeCount = basicGeometryShapeCount;
            MasterBackedShapeCount = masterBackedShapeCount;
            PackageBackedShapeCount = packageBackedShapeCount;
            GeneratedMasterBackedShapeCount = generatedMasterBackedShapeCount;
            SemanticOnlyShapeCount = semanticOnlyShapeCount;
        }

        /// <summary>Distinct stencil catalogs represented in the diagram stencil profile.</summary>
        public IReadOnlyList<string> StencilCatalogs { get; }

        /// <summary>Distinct Shape Data keys represented by shapes in the diagram stencil profile.</summary>
        public IReadOnlyList<string> ShapeDataKeys { get; }

        /// <summary>Distinct Shape Data keys represented by connectors in the diagram stencil profile.</summary>
        public IReadOnlyList<string> ConnectorShapeDataKeys { get; }

        /// <summary>Distinct OfficeIMO semantic kind values represented in the diagram stencil profile.</summary>
        public IReadOnlyList<string> SemanticKinds { get; }

        /// <summary>Number of distinct stencil catalogs represented in the diagram stencil profile.</summary>
        public int StencilCatalogCount => StencilCatalogs.Count;

        /// <summary>Number of distinct shape Shape Data keys represented in the diagram stencil profile.</summary>
        public int ShapeDataKeyCount => ShapeDataKeys.Count;

        /// <summary>Number of distinct connector Shape Data keys represented in the diagram stencil profile.</summary>
        public int ConnectorShapeDataKeyCount => ConnectorShapeDataKeys.Count;

        /// <summary>Number of distinct semantic kind values represented in the diagram stencil profile.</summary>
        public int SemanticKindCount => SemanticKinds.Count;

        /// <summary>Number of grouped stencil/profile usages reported by the diagram stencil profile.</summary>
        public int StencilUsageCount { get; }

        /// <summary>Total shape count reported by the diagram stencil profile.</summary>
        public int TotalShapeCount { get; }

        /// <summary>Total connector count reported by the diagram stencil profile.</summary>
        public int ConnectorCount { get; }

        /// <summary>Total connection points reported by the diagram stencil profile.</summary>
        public int TotalConnectionPointCount { get; }

        /// <summary>Number of shapes with connection points reported by the diagram stencil profile.</summary>
        public int ConnectionPointShapeCount { get; }

        /// <summary>Number of stencil families represented in the diagram stencil profile.</summary>
        public int StencilFamilyCount { get; }

        /// <summary>Number of shapes carrying OfficeIMO stencil metadata in the diagram stencil profile.</summary>
        public int StencilBackedShapeCount { get; }

        /// <summary>Number of basic geometry shapes reported by the diagram stencil profile.</summary>
        public int BasicGeometryShapeCount { get; }

        /// <summary>Number of master-backed shapes reported by the diagram stencil profile.</summary>
        public int MasterBackedShapeCount { get; }

        /// <summary>Number of package-backed shapes reported by the diagram stencil profile.</summary>
        public int PackageBackedShapeCount { get; }

        /// <summary>Number of generated-master-backed shapes reported by the diagram stencil profile.</summary>
        public int GeneratedMasterBackedShapeCount { get; }

        /// <summary>Number of semantic-only shapes reported by the diagram stencil profile.</summary>
        public int SemanticOnlyShapeCount { get; }
    }
}
