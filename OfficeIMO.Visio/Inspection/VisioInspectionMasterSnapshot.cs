using System.Collections.Generic;

namespace OfficeIMO.Visio {
/// <summary>
    /// Snapshot of a registered Visio master.
    /// </summary>
    public sealed class VisioInspectionMasterSnapshot {
        internal VisioInspectionMasterSnapshot(
            string id,
            string nameU,
            string? shapeNameU,
            string? text,
            double width,
            double height,
            bool isPackageBacked,
            string? stencilId,
            string? stencilName,
            string? stencilCategory,
            string? stencilCatalogName,
            string? stencilSourcePackagePath,
            IReadOnlyList<string> stencilKeywords,
            IReadOnlyList<string> stencilAliases,
            IReadOnlyList<string> stencilTags,
            string? stencilIconNameU,
            double? stencilDefaultWidth,
            double? stencilDefaultHeight,
            string? stencilDefaultUnit,
            string? stencilPreviewImageRelationshipId,
            string? stencilPreviewImageTarget,
            string? stencilPreviewImageContentType,
            string? stencilPreviewImageExtension,
            long? stencilPreviewImageByteLength) {
            Id = id;
            NameU = nameU;
            ShapeNameU = shapeNameU;
            Text = text;
            Width = width;
            Height = height;
            IsPackageBacked = isPackageBacked;
            StencilId = stencilId;
            StencilName = stencilName;
            StencilCategory = stencilCategory;
            StencilCatalogName = stencilCatalogName;
            StencilSourcePackagePath = stencilSourcePackagePath;
            StencilKeywords = stencilKeywords;
            StencilAliases = stencilAliases;
            StencilTags = stencilTags;
            StencilIconNameU = stencilIconNameU;
            StencilDefaultWidth = stencilDefaultWidth;
            StencilDefaultHeight = stencilDefaultHeight;
            StencilDefaultUnit = stencilDefaultUnit;
            StencilPreviewImageRelationshipId = stencilPreviewImageRelationshipId;
            StencilPreviewImageTarget = stencilPreviewImageTarget;
            StencilPreviewImageContentType = stencilPreviewImageContentType;
            StencilPreviewImageExtension = stencilPreviewImageExtension;
            StencilPreviewImageByteLength = stencilPreviewImageByteLength;
        }

        /// <summary>Master identifier.</summary>
        public string Id { get; }

        /// <summary>Master universal name.</summary>
        public string NameU { get; }

        /// <summary>Universal name of the master shape.</summary>
        public string? ShapeNameU { get; }

        /// <summary>Text stored on the master shape.</summary>
        public string? Text { get; }

        /// <summary>Master shape width.</summary>
        public double Width { get; }

        /// <summary>Master shape height.</summary>
        public double Height { get; }

        /// <summary>Whether the master came from a package-backed stencil or document.</summary>
        public bool IsPackageBacked { get; }

        /// <summary>OfficeIMO stencil identifier, when known.</summary>
        public string? StencilId { get; }

        /// <summary>OfficeIMO stencil display name, when known.</summary>
        public string? StencilName { get; }

        /// <summary>OfficeIMO stencil category, when known.</summary>
        public string? StencilCategory { get; }

        /// <summary>Stencil catalog name, when known.</summary>
        public string? StencilCatalogName { get; }

        /// <summary>Source package path, when known.</summary>
        public string? StencilSourcePackagePath { get; }

        /// <summary>Searchable stencil keywords.</summary>
        public IReadOnlyList<string> StencilKeywords { get; }

        /// <summary>Stencil lookup aliases.</summary>
        public IReadOnlyList<string> StencilAliases { get; }

        /// <summary>Semantic stencil tags.</summary>
        public IReadOnlyList<string> StencilTags { get; }

        /// <summary>Preview icon master universal name, when known.</summary>
        public string? StencilIconNameU { get; }

        /// <summary>Source stencil default width, when known.</summary>
        public double? StencilDefaultWidth { get; }

        /// <summary>Source stencil default height, when known.</summary>
        public double? StencilDefaultHeight { get; }

        /// <summary>Source stencil default size unit, when known.</summary>
        public string? StencilDefaultUnit { get; }

        /// <summary>Preview image relationship id, when known.</summary>
        public string? StencilPreviewImageRelationshipId { get; }

        /// <summary>Preview image relationship target, when known.</summary>
        public string? StencilPreviewImageTarget { get; }

        /// <summary>Preview image content type, when known.</summary>
        public string? StencilPreviewImageContentType { get; }

        /// <summary>Preview image extension, when known.</summary>
        public string? StencilPreviewImageExtension { get; }

        /// <summary>Preview image byte length, when known.</summary>
        public long? StencilPreviewImageByteLength { get; }
    }
}
