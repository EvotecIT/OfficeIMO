using System;
using System.IO;
using System.Linq;

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
            string contentType = ResolveContentType(relationship);
            return contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase) ||
                   IsImageExtension(relationship.Extension) ||
                   IsImageExtension(Path.GetExtension(relationship.Target));
        }

        private static string ResolveContentType(VisioAssets.MasterRelationshipContent relationship) {
            string normalizedContentType = NormalizeContentType(relationship.ContentType);
            if (!string.IsNullOrWhiteSpace(normalizedContentType) &&
                !IsGenericContentType(normalizedContentType)) {
                return normalizedContentType;
            }

            if (TrySniffContentType(relationship.Data, out string? sniffedContentType)) {
                return sniffedContentType!;
            }

            string extension = string.IsNullOrWhiteSpace(relationship.Extension)
                ? Path.GetExtension(relationship.Target)
                : relationship.Extension;
            switch (extension?.ToLowerInvariant()) {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".svg":
                    return "image/svg+xml";
                default:
                    return "application/octet-stream";
            }
        }

        private static bool IsBrowserRenderable(string contentType, string? extension) {
            if (string.Equals(contentType, "image/png", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(contentType, "image/jpeg", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(contentType, "image/gif", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(contentType, "image/svg+xml", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            return IsImageExtension(extension);
        }

        private static string NormalizeContentType(string? contentType) {
            if (string.IsNullOrWhiteSpace(contentType)) {
                return string.Empty;
            }

            int separator = contentType!.IndexOf(';');
            string normalized = separator >= 0
                ? contentType.Substring(0, separator)
                : contentType;
            return normalized.Trim().ToLowerInvariant();
        }

        private static bool IsGenericContentType(string contentType) =>
            string.Equals(contentType, "application/octet-stream", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(contentType, "binary/octet-stream", StringComparison.OrdinalIgnoreCase);

        private static bool TrySniffContentType(byte[]? data, out string? contentType) {
            contentType = null;
            if (data == null || data.Length == 0) {
                return false;
            }

            if (data.Length >= 8 &&
                data[0] == 0x89 &&
                data[1] == (byte)'P' &&
                data[2] == (byte)'N' &&
                data[3] == (byte)'G' &&
                data[4] == 0x0D &&
                data[5] == 0x0A &&
                data[6] == 0x1A &&
                data[7] == 0x0A) {
                contentType = "image/png";
                return true;
            }

            if (data.Length >= 3 &&
                data[0] == 0xFF &&
                data[1] == 0xD8 &&
                data[2] == 0xFF) {
                contentType = "image/jpeg";
                return true;
            }

            if (data.Length >= 6 &&
                data[0] == (byte)'G' &&
                data[1] == (byte)'I' &&
                data[2] == (byte)'F' &&
                data[3] == (byte)'8' &&
                (data[4] == (byte)'7' || data[4] == (byte)'9') &&
                data[5] == (byte)'a') {
                contentType = "image/gif";
                return true;
            }

            if (LooksLikeSvg(data)) {
                contentType = "image/svg+xml";
                return true;
            }

            return false;
        }

        private static bool LooksLikeSvg(byte[] data) {
            int index = SkipBomAndWhitespace(data, 0);
            while (index < data.Length && data[index] == (byte)'<') {
                int tagStart = SkipAsciiWhitespace(data, index + 1);
                if (StartsWithAscii(data, tagStart, "svg")) {
                    return true;
                }

                if (StartsWithAscii(data, tagStart, "!--")) {
                    int commentEnd = IndexOfAscii(data, tagStart + 3, "-->");
                    if (commentEnd < 0) {
                        return false;
                    }

                    index = SkipAsciiWhitespace(data, commentEnd + 3);
                    continue;
                }

                if (StartsWithAscii(data, tagStart, "!doctype")) {
                    int declarationEnd = IndexOfByte(data, tagStart + 8, (byte)'>');
                    if (declarationEnd < 0) {
                        return false;
                    }

                    index = SkipAsciiWhitespace(data, declarationEnd + 1);
                    continue;
                }

                if (tagStart < data.Length && data[tagStart] == (byte)'?') {
                    int processingInstructionEnd = IndexOfAscii(data, tagStart + 1, "?>");
                    if (processingInstructionEnd < 0) {
                        return false;
                    }

                    index = SkipAsciiWhitespace(data, processingInstructionEnd + 2);
                    continue;
                }

                return false;
            }

            return false;
        }

        private static int SkipBomAndWhitespace(byte[] data, int index) {
            if (data.Length >= index + 3 &&
                data[index] == 0xEF &&
                data[index + 1] == 0xBB &&
                data[index + 2] == 0xBF) {
                index += 3;
            }

            return SkipAsciiWhitespace(data, index);
        }

        private static int SkipAsciiWhitespace(byte[] data, int index) {
            while (index < data.Length && IsAsciiWhitespace(data[index])) {
                index++;
            }

            return index;
        }

        private static int IndexOfByte(byte[] data, int startIndex, byte value) {
            for (int i = startIndex; i < data.Length; i++) {
                if (data[i] == value) {
                    return i;
                }
            }

            return -1;
        }

        private static bool IsAsciiWhitespace(byte value) =>
            value == (byte)' ' ||
            value == (byte)'\t' ||
            value == (byte)'\r' ||
            value == (byte)'\n';

        private static int IndexOfAscii(byte[] data, int startIndex, string value) {
            for (int i = startIndex; i <= data.Length - value.Length; i++) {
                if (StartsWithAscii(data, i, value)) {
                    return i;
                }
            }

            return -1;
        }

        private static bool StartsWithAscii(byte[] data, int startIndex, string value) {
            if (startIndex < 0 || startIndex + value.Length > data.Length) {
                return false;
            }

            for (int i = 0; i < value.Length; i++) {
                byte actual = data[startIndex + i];
                byte expected = (byte)value[i];
                if (actual >= (byte)'A' && actual <= (byte)'Z') {
                    actual = (byte)(actual + 32);
                }

                if (expected >= (byte)'A' && expected <= (byte)'Z') {
                    expected = (byte)(expected + 32);
                }

                if (actual != expected) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsImageExtension(string? extension) {
            switch (extension?.ToLowerInvariant()) {
                case ".png":
                case ".jpg":
                case ".jpeg":
                case ".gif":
                case ".svg":
                    return true;
                default:
                    return false;
            }
        }

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
