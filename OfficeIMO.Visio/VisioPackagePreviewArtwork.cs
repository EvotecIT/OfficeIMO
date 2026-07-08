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

        internal static bool TryGetRasterImage(VisioShape shape, out OfficeRasterImage? image) {
            image = null;
            if (!TryGetPreviewRelationship(shape, out VisioAssets.MasterRelationshipContent? relationship) || relationship == null) {
                return false;
            }

            if (OfficeRasterImageDecoder.TryDecode(relationship.Data, out OfficeRasterImage? raster) && raster != null) {
                image = raster;
                return true;
            }

            if (IsSvgRelationship(relationship) &&
                VisioSvgPreviewRasterizer.TryRasterize(relationship.Data, href => TryResolveRelatedImage(shape, relationship, href), out OfficeRasterImage? svgRaster) &&
                svgRaster != null) {
                image = svgRaster;
                return true;
            }

            return false;
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
                   OfficeImageInfo.IsBrowserPreviewSafeExtension(Path.GetExtension(relationship.Target)) ||
                   OfficeRasterImageDecoder.TryDecode(relationship.Data, out _);
        }

        private static string ResolveContentType(VisioAssets.MasterRelationshipContent relationship) {
            return OfficeSvgImageRenderer.TryResolveEmbeddableContentType(
                relationship.ContentType,
                relationship.Data,
                GetRelationshipImageName(relationship),
                out string contentType)
                ? contentType
                : OfficeImageInfo.GetMimeType(OfficeImageFormat.Unknown);
        }

        private static bool IsBrowserRenderable(string contentType, string? extension) {
            if (OfficeSvgImageRenderer.TryGetEmbeddableContentType(contentType, out _)) {
                return true;
            }

            return OfficeImageInfo.IsBrowserPreviewSafeExtension(extension);
        }

        private static bool IsSvgRelationship(VisioAssets.MasterRelationshipContent relationship) =>
            OfficeImageInfo.FromMimeType(relationship.ContentType) == OfficeImageFormat.Svg ||
            OfficeImageReader.FromExtension(GetRelationshipImageName(relationship)) == OfficeImageFormat.Svg;

        private static byte[]? TryResolveRelatedImage(VisioShape shape, VisioAssets.MasterRelationshipContent svgRelationship, string href) {
            if (shape.Master?.RawMasterRelationships.Count > 0 != true || string.IsNullOrWhiteSpace(href)) {
                return null;
            }

            string normalizedHref = NormalizePath(href);
            string svgDirectory = GetDirectoryName(NormalizePath(svgRelationship.Target));
            string relativeToSvg = string.IsNullOrWhiteSpace(svgDirectory)
                ? normalizedHref
                : NormalizePath(svgDirectory + "/" + normalizedHref);
            string hrefFileName = Path.GetFileName(normalizedHref);

            VisioAssets.MasterRelationshipContent? match = shape.Master.RawMasterRelationships
                .Where(item => !item.IsExternal && item.Data != null && item.Data.Length > 0 && !ReferenceEquals(item, svgRelationship))
                .Where(IsImageRelationship)
                .FirstOrDefault(item => MatchesImageHref(item, normalizedHref, relativeToSvg, hrefFileName));

            return match?.Data;
        }

        private static bool MatchesImageHref(VisioAssets.MasterRelationshipContent relationship, string normalizedHref, string relativeToSvg, string hrefFileName) {
            string target = NormalizePath(relationship.Target);
            if (string.Equals(target, normalizedHref, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(target, relativeToSvg, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return hrefFileName.Length > 0 &&
                   string.Equals(Path.GetFileName(target), hrefFileName, StringComparison.OrdinalIgnoreCase);
        }

        private static string GetDirectoryName(string normalizedTarget) {
            int slash = normalizedTarget.LastIndexOf('/');
            return slash <= 0 ? string.Empty : normalizedTarget.Substring(0, slash);
        }

        private static string GetRelationshipImageName(VisioAssets.MasterRelationshipContent relationship) =>
            string.IsNullOrWhiteSpace(relationship.Extension)
                ? Path.GetExtension(relationship.Target)
                : relationship.Extension;

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
