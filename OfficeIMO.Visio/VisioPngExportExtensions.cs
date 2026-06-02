using System;
using System.IO;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Native dependency-free PNG export helpers for OfficeIMO Visio documents and pages.
    /// </summary>
    public static class VisioPngExportExtensions {
        /// <summary>
        /// Renders the selected document page to PNG bytes without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static byte[] ToPng(this VisioDocument document, VisioPngSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new VisioPngSaveOptions();
            if (document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages to export.");
            }

            if (options.PageIndex < 0 || options.PageIndex >= document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
            }

            return VisioPngRenderer.Render(document.Pages[options.PageIndex], options);
        }

        /// <summary>
        /// Renders a page to PNG bytes without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static byte[] ToPng(this VisioPage page, VisioPngSaveOptions? options = null) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            return VisioPngRenderer.Render(page, options ?? new VisioPngSaveOptions());
        }

        /// <summary>
        /// Saves the selected document page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsPng(this VisioDocument document, string path, VisioPngSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            WritePng(path, document.ToPng(options));
        }

        /// <summary>
        /// Saves a page as PNG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsPng(this VisioPage page, string path, VisioPngSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("PNG path cannot be null or whitespace.", nameof(path));
            WritePng(path, page.ToPng(options));
        }

        /// <summary>
        /// Writes the selected document page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsPng(this VisioDocument document, Stream stream, VisioPngSaveOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            byte[] bytes = document.ToPng(options);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// Writes a page as PNG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsPng(this VisioPage page, Stream stream, VisioPngSaveOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            byte[] bytes = page.ToPng(options);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WritePng(string path, byte[] bytes) {
            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            File.WriteAllBytes(fullPath, bytes);
        }
    }
}
