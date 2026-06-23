using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio {
    internal static class VisioPackagePreviewArtwork {
        internal static bool TryGetBrowserImage(VisioShape shape, out VisioPreviewImage image) {
            image = default;
            if (!TryGetPreviewRelationship(shape, out VisioAssets.MasterRelationshipContent? relationship) || relationship == null) {
                return false;
            }

            string contentType = ResolveContentType(relationship);
            if (!IsBrowserRenderable(contentType, relationship.Extension)) {
                return false;
            }

            image = new VisioPreviewImage(contentType, relationship.Data!);
            return true;
        }

        internal static bool TryGetPng(VisioShape shape, out VisioPreviewImage image) {
            image = default;
            if (!TryGetPreviewRelationship(shape, out VisioAssets.MasterRelationshipContent? relationship) || relationship == null) {
                return false;
            }

            string contentType = ResolveContentType(relationship);
            if (!string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(relationship.Extension, ".png", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            image = new VisioPreviewImage("image/png", relationship.Data!);
            return true;
        }

        internal static bool HasPreviewMetadata(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            string? relationshipId = shape.GetUserCellValue(VisioSemanticUserCells.StencilPreviewImageRelationshipId) ??
                                     shape.Master?.StencilPreviewImageRelationshipId;
            string? target = shape.GetUserCellValue(VisioSemanticUserCells.StencilPreviewImageTarget) ??
                             shape.Master?.StencilPreviewImageTarget;
            return !string.IsNullOrWhiteSpace(relationshipId) ||
                   !string.IsNullOrWhiteSpace(target);
        }

        private static bool TryGetPreviewRelationship(VisioShape shape, out VisioAssets.MasterRelationshipContent? relationship) {
            relationship = null;
            if (shape.Master?.RawMasterRelationships.Count > 0 != true) {
                return false;
            }

            string? relationshipId = shape.GetUserCellValue(VisioSemanticUserCells.StencilPreviewImageRelationshipId) ??
                                     shape.Master.StencilPreviewImageRelationshipId;
            string? target = shape.GetUserCellValue(VisioSemanticUserCells.StencilPreviewImageTarget) ??
                             shape.Master.StencilPreviewImageTarget;

            bool hasRelationshipId = !string.IsNullOrWhiteSpace(relationshipId);
            bool hasTarget = !string.IsNullOrWhiteSpace(target);
            string? normalizedTarget = hasTarget ? NormalizePath(target!) : null;

            relationship = shape.Master.RawMasterRelationships
                .Where(item => !item.IsExternal && item.Data != null && item.Data.Length > 0)
                .Where(item => !hasRelationshipId || string.Equals(item.Id, relationshipId, StringComparison.OrdinalIgnoreCase))
                .Where(item => !hasTarget || string.Equals(NormalizePath(item.Target), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                .OrderBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(IsImageRelationship);

            return relationship != null;
        }

        private static bool IsImageRelationship(VisioAssets.MasterRelationshipContent relationship) {
            return OfficeSvgImageRenderer.TryResolveEmbeddableContentType(
                       relationship.ContentType,
                       relationship.Data,
                       GetRelationshipImageName(relationship),
                       out _) ||
                   IsBrowserRenderableExtension(Path.GetExtension(relationship.Target));
        }

        private static string ResolveContentType(VisioAssets.MasterRelationshipContent relationship) {
            return OfficeSvgImageRenderer.TryResolveEmbeddableContentType(
                relationship.ContentType,
                relationship.Data,
                GetRelationshipImageName(relationship),
                out string contentType)
                ? contentType
                : "application/octet-stream";
        }

        private static bool IsBrowserRenderable(string contentType, string? extension) {
            if (string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(contentType, "image/gif", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return IsBrowserRenderableExtension(extension);
        }

        private static string GetRelationshipImageName(VisioAssets.MasterRelationshipContent relationship) =>
            string.IsNullOrWhiteSpace(relationship.Extension)
                ? Path.GetExtension(relationship.Target)
                : relationship.Extension;

        private static bool IsBrowserRenderableExtension(string? extension) =>
            OfficeSvgImageRenderer.TryGetEmbeddableContentType(OfficeImageReader.FromExtension(extension), out _);

        private static string NormalizePath(string value) =>
            value.Replace('\\', '/').TrimStart('/');
    }

    internal readonly struct VisioPreviewImage {
        internal VisioPreviewImage(string contentType, byte[] data) {
            ContentType = contentType;
            Data = data;
        }

        internal string ContentType { get; }

        internal byte[] Data { get; }
    }
}
