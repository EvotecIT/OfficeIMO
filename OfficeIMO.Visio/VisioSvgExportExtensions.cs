using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Headless SVG export helpers for OfficeIMO Visio documents and pages.
    /// </summary>
    public static class VisioSvgExportExtensions {
        /// <summary>
        /// Renders the selected document page to SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static string ToSvg(this VisioDocument document, VisioSvgSaveOptions? options = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new VisioSvgSaveOptions();
            if (document.Pages.Count == 0) {
                throw new InvalidOperationException("The document does not contain any pages to export.");
            }

            if (options.PageIndex < 0 || options.PageIndex >= document.Pages.Count) {
                throw new ArgumentOutOfRangeException(nameof(options), "PageIndex is outside the document page collection.");
            }

            return VisioSvgRenderer.Render(document.Pages[options.PageIndex], options);
        }

        /// <summary>
        /// Renders a page to SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static string ToSvg(this VisioPage page, VisioSvgSaveOptions? options = null) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            return VisioSvgRenderer.Render(page, options ?? new VisioSvgSaveOptions());
        }

        /// <summary>
        /// Saves the selected document page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsSvg(this VisioDocument document, string path, VisioSvgSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("SVG path cannot be null or whitespace.", nameof(path));
            string svg = document.ToSvg(options);
            WriteSvg(path, svg);
        }

        /// <summary>
        /// Saves a page as SVG without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsSvg(this VisioPage page, string path, VisioSvgSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("SVG path cannot be null or whitespace.", nameof(path));
            string svg = page.ToSvg(options);
            WriteSvg(path, svg);
        }

        /// <summary>
        /// Writes the selected document page as SVG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsSvg(this VisioDocument document, Stream stream, VisioSvgSaveOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            WriteSvg(stream, document.ToSvg(options));
        }

        /// <summary>
        /// Writes a page as SVG to a stream without requiring Microsoft Visio desktop automation.
        /// </summary>
        public static void SaveAsSvg(this VisioPage page, Stream stream, VisioSvgSaveOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            WriteSvg(stream, page.ToSvg(options));
        }

        private static void WriteSvg(string path, string svg) {
            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            File.WriteAllText(fullPath, svg, Encoding.UTF8);
        }

        private static void WriteSvg(Stream stream, string svg) {
            byte[] bytes = Encoding.UTF8.GetBytes(svg);
            stream.Write(bytes, 0, bytes.Length);
        }
    }
}
