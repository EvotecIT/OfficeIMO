using System;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this slide.
        /// </summary>
        public PowerPointSlideVisualSnapshot CreateVisualSnapshot(PowerPointImageExportOptions? options = null) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointSlideImageRenderer.CreateSnapshot(this, resolved);
        }

        /// <summary>
        /// Exports this slide as PNG or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, PowerPointImageExportOptions? options = null) {
            PowerPointImageExportOptions resolved = NormalizeImageExportOptions(options);
            return PowerPointSlideImageRenderer.Render(this, format, resolved);
        }

        /// <summary>
        /// Renders this slide to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(PowerPointImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders this slide to dependency-free SVG text.
        /// </summary>
        public string ToSvg(PowerPointImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves this slide as a PNG file.
        /// </summary>
        public void SaveAsPng(string path, PowerPointImageExportOptions? options = null) =>
            WriteImageFile(path, ToPng(options));

        /// <summary>
        /// Saves this slide as an SVG file.
        /// </summary>
        public void SaveAsSvg(string path, PowerPointImageExportOptions? options = null) =>
            WriteImageFile(path, Encoding.UTF8.GetBytes(ToSvg(options)));

        /// <summary>
        /// Writes this slide as PNG to a stream.
        /// </summary>
        public void SaveAsPng(Stream stream, PowerPointImageExportOptions? options = null) =>
            WriteImageStream(stream, ToPng(options));

        /// <summary>
        /// Writes this slide as SVG to a stream.
        /// </summary>
        public void SaveAsSvg(Stream stream, PowerPointImageExportOptions? options = null) =>
            WriteImageStream(stream, Encoding.UTF8.GetBytes(ToSvg(options)));

        private static PowerPointImageExportOptions NormalizeImageExportOptions(PowerPointImageExportOptions? options) {
            PowerPointImageExportOptions resolved = options?.Clone() ?? new PowerPointImageExportOptions();
            OfficeImageExportOptions.ValidateScale(resolved.Scale, nameof(options));
            return resolved;
        }

        private static void WriteImageFile(string path, byte[] bytes) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
            }

            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            OfficeIMO.Core.Internal.OfficeFileCommit.WriteAllBytes(fullPath, bytes);
        }

        private static void WriteImageStream(Stream stream, byte[] bytes) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            stream.Write(bytes, 0, bytes.Length);
        }
    }
}
