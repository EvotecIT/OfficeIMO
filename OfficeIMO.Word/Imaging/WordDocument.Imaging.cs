using System;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>
        /// Creates a format-neutral visual snapshot for the requested document page.
        /// </summary>
        public WordDocumentVisualSnapshot CreateVisualSnapshot(WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.CreateSnapshot(this, resolved);
        }

        /// <summary>
        /// Exports the requested document page as PNG or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, WordImageExportOptions? options = null) {
            WordImageExportOptions resolved = NormalizeImageExportOptions(options);
            return WordDocumentImageRenderer.Render(this, format, resolved);
        }

        /// <summary>
        /// Renders the requested document page to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(WordImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders the requested document page to dependency-free SVG text.
        /// </summary>
        public string ToSvg(WordImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves the requested document page as a PNG file.
        /// </summary>
        public void SaveAsPng(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().Save(path);

        /// <summary>
        /// Saves the requested document page as an SVG file.
        /// </summary>
        public void SaveAsSvg(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().Save(path);

        /// <summary>
        /// Writes the requested document page as PNG to a stream.
        /// </summary>
        public void SaveAsPng(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsPng().Save(stream);

        /// <summary>
        /// Writes the requested document page as SVG to a stream.
        /// </summary>
        public void SaveAsSvg(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsSvg().Save(stream);

        private static WordImageExportOptions NormalizeImageExportOptions(WordImageExportOptions? options) {
            WordImageExportOptions resolved = options?.Clone() ?? new WordImageExportOptions();
            OfficeImageExportOptions.ValidateScale(resolved.Scale, nameof(options));
            if (resolved.PageIndex < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Page index cannot be negative.");
            }

            return resolved;
        }
    }
}
