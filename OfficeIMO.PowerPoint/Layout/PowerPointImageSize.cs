using System;
using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>Reads the pixel size of an image so pictures keep their aspect ratio.</summary>
    internal static class PowerPointImageSize {
        internal static bool TryRead(string path, out int width, out int height) {
            width = 0;
            height = 0;
            try {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
                using FileStream stream = File.OpenRead(path);
                return TryRead(stream, path, out width, out height);
            } catch (IOException) {
                return false;
            } catch (UnauthorizedAccessException) {
                return false;
            }
        }

        internal static bool TryRead(Stream stream, string? name, out int width, out int height) {
            width = 0;
            height = 0;
            long position = stream.CanSeek ? stream.Position : 0;
            try {
                if (!OfficeImageReader.TryIdentify(stream, name, out OfficeImageInfo info)) return false;
                width = info.Width;
                height = info.Height;
                return width > 0 && height > 0;
            } finally {
                if (stream.CanSeek) stream.Position = position;
            }
        }

        internal static bool TryRead(byte[] data, out int width, out int height) {
            width = 0;
            height = 0;
            if (!OfficeImageReader.TryIdentify(data, null, out OfficeImageInfo info)) return false;
            width = info.Width;
            height = info.Height;
            return width > 0 && height > 0;
        }
    }
}
