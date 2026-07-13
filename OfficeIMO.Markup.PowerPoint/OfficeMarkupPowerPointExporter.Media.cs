using System.Diagnostics;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Markup.PowerPoint;

internal sealed partial class OfficeMarkupPowerPointExporter {
    private static void AddImage(
        PowerPointSlide slide,
        OfficeMarkupImageBlock image,
        LayoutCursor cursor,
        MarkupToPowerPointOptions options,
        SlideCanvasMetrics metrics) {
        if (TryResolveFilePath(options, image.Source, out var path) && File.Exists(path)) {
            var box = ResolveBox(image.Placement, image.Attributes, cursor, Math.Min(2.2, cursor.RemainingHeight), metrics);
            if (ShouldAddVisualPanel(image.Attributes, defaultValue: false)) {
                AddVisualPanel(slide, box, metrics, "OfficeIMO Markup Image Panel");
            }

            AddPicture(slide, path, box, GetAttribute(image.Attributes, "fit"));
            if (!HasExplicitPlacement(image.Placement, image.Attributes)) {
                cursor.Advance(box.Height);
            }
        } else if (options.IncludeUnsupportedBlocksAsText) {
            AddText(slide, $"Image: {image.Source}", cursor, height: 0.4);
        }
    }

    private static bool TryResolveFilePath(
        MarkupToPowerPointOptions? options,
        string source,
        out string path) {
        path = source;
        if (string.IsNullOrWhiteSpace(source)) {
            return false;
        }

        if (Uri.TryCreate(source, UriKind.Absolute, out var uri)) {
            if (!uri.IsFile) {
                return false;
            }

            path = uri.LocalPath;
        } else {
            path = source;
        }

        try {
            bool hasBaseDirectory = options != null && !string.IsNullOrWhiteSpace(options.BaseDirectory);
            if (!hasBaseDirectory && !(options?.AllowExternalImagePaths ?? false)) {
                return false;
            }

            string candidate = Path.IsPathRooted(path) || !hasBaseDirectory
                ? path
                : Path.Combine(options!.BaseDirectory!, path);
            path = Path.GetFullPath(candidate);
            if (hasBaseDirectory
                && !(options?.AllowExternalImagePaths ?? false)
                && !IsPathWithinDirectory(path, Path.GetFullPath(options!.BaseDirectory!))) {
                return false;
            }

            return true;
        } catch (Exception) when (!Debugger.IsAttached) {
            return false;
        }
    }

    private static bool IsPathWithinDirectory(string path, string directory) {
        string normalizedDirectory = directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        string normalizedPath = path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        return normalizedPath.StartsWith(normalizedDirectory, StringComparison.OrdinalIgnoreCase);
    }

    private static void AddTable(PowerPointSlide slide, OfficeMarkupTableBlock table, LayoutCursor cursor) {
        var rows = new List<IReadOnlyList<string>>();
        if (table.Headers.Count > 0) {
            rows.Add(table.Headers.ToList());
        }

        rows.AddRange(table.Rows);
        if (rows.Count == 0) {
            return;
        }

        var columns = rows.Max(row => row.Count);
        var height = Math.Min(cursor.RemainingHeight, Math.Max(0.8, 0.32 * rows.Count));
        var powerPointTable = slide.AddTableInches(rows.Count, Math.Max(1, columns), cursor.Left, cursor.Top, cursor.Width, height);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            for (int columnIndex = 0; columnIndex < row.Count; columnIndex++) {
                powerPointTable.GetCell(rowIndex, columnIndex).Text = row[columnIndex];
            }
        }

        cursor.Advance(height);
    }

    private static bool ShouldAddVisualPanel(IDictionary<string, string> attributes, bool defaultValue) {
        var value = GetAttribute(attributes, "panel", "frame", "background-panel");
        return string.IsNullOrWhiteSpace(value) ? defaultValue : IsTruthy(value!);
    }

    private static void AddVisualPanel(PowerPointSlide slide, LayoutCursor box, SlideCanvasMetrics metrics, string name) {
        const double padding = 0.08;
        var left = Math.Max(metrics.Horizontal(0.18), box.Left - metrics.Horizontal(padding));
        var top = Math.Max(metrics.Vertical(0.18), box.Top - metrics.Vertical(padding));
        var right = Math.Min(metrics.Width - metrics.Horizontal(0.18), box.Left + box.Width + metrics.Horizontal(padding));
        var bottom = Math.Min(metrics.Height - metrics.Vertical(0.18), box.Top + box.Height + metrics.Vertical(padding));
        var panel = slide.AddShapeInches(
            A.ShapeTypeValues.Rectangle,
            left,
            top,
            Math.Max(0.5, right - left),
            Math.Max(0.5, bottom - top),
            name);
        panel.FillColor = "FFFFFF";
        panel.FillTransparency = 4;
        panel.OutlineColor = "D9E2EF";
        panel.OutlineWidthPoints = 0.75;
    }

    private static void AddPicture(PowerPointSlide slide, string path, LayoutCursor box, string? fit) {
        switch (Normalize(fit ?? string.Empty)) {
            case "fill":
            case "stretch":
                slide.AddPictureInches(path, box.Left, box.Top, box.Width, box.Height);
                return;
            case "contain":
            default:
                AddPictureContained(slide, path, box);
                return;
        }
    }

    private static void AddPictureContained(PowerPointSlide slide, string path, LayoutCursor box) {
        var left = box.Left;
        var top = box.Top;
        var width = box.Width;
        var height = box.Height;

        if (TryReadImageSize(path, out var pixelWidth, out var pixelHeight) && pixelWidth > 0 && pixelHeight > 0) {
            var imageAspect = pixelWidth / (double)pixelHeight;
            var boxAspect = box.Width / box.Height;
            if (imageAspect > boxAspect) {
                height = box.Width / imageAspect;
                top = box.Top + ((box.Height - height) / 2.0);
            } else {
                width = box.Height * imageAspect;
                left = box.Left + ((box.Width - width) / 2.0);
            }
        }

        slide.AddPictureInches(path, left, top, width, height);
    }

    private static bool TryReadImageSize(string path, out int width, out int height) {
        width = 0;
        height = 0;

        try {
            using var stream = File.OpenRead(path);
            if (TryReadPngSize(stream, out width, out height)) {
                return true;
            }

            stream.Position = 0;
            return TryReadJpegSize(stream, out width, out height);
        } catch (IOException) {
            return false;
        } catch (UnauthorizedAccessException) {
            return false;
        }
    }

    private static bool TryReadPngSize(Stream stream, out int width, out int height) {
        width = 0;
        height = 0;

        var header = new byte[24];
        if (stream.Read(header, 0, header.Length) != header.Length || !IsPngHeader(header)) {
            return false;
        }

        width = ReadBigEndianInt32(header, 16);
        height = ReadBigEndianInt32(header, 20);
        return true;
    }

    private static bool TryReadJpegSize(Stream stream, out int width, out int height) {
        width = 0;
        height = 0;

        if (stream.ReadByte() != 0xFF || stream.ReadByte() != 0xD8) {
            return false;
        }

        while (stream.Position < stream.Length) {
            int prefix;
            do {
                prefix = stream.ReadByte();
            } while (prefix != -1 && prefix != 0xFF);

            if (prefix == -1) {
                return false;
            }

            int marker;
            do {
                marker = stream.ReadByte();
            } while (marker == 0xFF);

            if (marker == -1) {
                return false;
            }

            if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01) {
                continue;
            }

            var segmentLength = ReadBigEndianUInt16(stream);
            if (segmentLength < 2) {
                return false;
            }

            if (IsJpegStartOfFrame(marker)) {
                if (segmentLength < 7) {
                    return false;
                }

                if (stream.ReadByte() == -1) {
                    return false;
                }

                height = ReadBigEndianUInt16(stream);
                width = ReadBigEndianUInt16(stream);
                return width > 0 && height > 0;
            }

            stream.Seek(segmentLength - 2, SeekOrigin.Current);
        }

        return false;
    }

    private static bool IsPngHeader(byte[] header) =>
        header.Length >= 24
        && header[0] == 0x89
        && header[1] == 0x50
        && header[2] == 0x4E
        && header[3] == 0x47
        && header[4] == 0x0D
        && header[5] == 0x0A
        && header[6] == 0x1A
        && header[7] == 0x0A;

    private static int ReadBigEndianInt32(byte[] value, int offset) =>
        (value[offset] << 24) | (value[offset + 1] << 16) | (value[offset + 2] << 8) | value[offset + 3];

    private static int ReadBigEndianUInt16(Stream stream) {
        var high = stream.ReadByte();
        var low = stream.ReadByte();
        return high < 0 || low < 0 ? -1 : (high << 8) | low;
    }

    private static bool IsJpegStartOfFrame(int marker) =>
        marker is 0xC0 or 0xC1 or 0xC2 or 0xC3 or 0xC5 or 0xC6 or 0xC7 or 0xC9 or 0xCA or 0xCB or 0xCD or 0xCE or 0xCF;
}
