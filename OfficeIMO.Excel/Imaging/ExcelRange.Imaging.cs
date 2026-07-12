using OfficeIMO.Drawing.Internal;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelRange {
        /// <summary>
        /// Creates a format-neutral visual snapshot for this range.
        /// </summary>
        public ExcelRangeVisualSnapshot CreateVisualSnapshot(ExcelImageExportOptions? options = null) =>
            ExcelRangeVisualSnapshotBuilder.Build(Sheet, Address, NormalizeOptions(options));

        /// <summary>
        /// Exports this range as PNG or SVG.
        /// </summary>
        public OfficeImageExportResult ExportImage(OfficeImageExportFormat format, ExcelImageExportOptions? options = null) {
            ExcelImageExportOptions resolved = NormalizeOptions(options);
            ExcelRangeVisualSnapshot snapshot = ExcelRangeVisualSnapshotBuilder.Build(Sheet, Address, resolved);
            return ExcelRangeImageRenderer.Render(snapshot, format, resolved);
        }

        /// <summary>
        /// Renders this range to dependency-free PNG bytes.
        /// </summary>
        public byte[] ToPng(ExcelImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Png, options).Bytes;

        /// <summary>
        /// Renders this range to dependency-free SVG text.
        /// </summary>
        public string ToSvg(ExcelImageExportOptions? options = null) =>
            Encoding.UTF8.GetString(ExportImage(OfficeImageExportFormat.Svg, options).Bytes);

        /// <summary>
        /// Saves this range as a PNG file.
        /// </summary>
        public void SaveAsPng(string path, ExcelImageExportOptions? options = null) =>
            WriteFile(path, ToPng(options));

        /// <summary>
        /// Saves this range as an SVG file.
        /// </summary>
        public void SaveAsSvg(string path, ExcelImageExportOptions? options = null) =>
            WriteFile(path, Encoding.UTF8.GetBytes(ToSvg(options)));

        /// <summary>
        /// Writes this range as PNG to a stream.
        /// </summary>
        public void SaveAsPng(Stream stream, ExcelImageExportOptions? options = null) =>
            WriteStream(stream, ToPng(options));

        /// <summary>
        /// Writes this range as SVG to a stream.
        /// </summary>
        public void SaveAsSvg(Stream stream, ExcelImageExportOptions? options = null) =>
            WriteStream(stream, Encoding.UTF8.GetBytes(ToSvg(options)));

        private static ExcelImageExportOptions NormalizeOptions(ExcelImageExportOptions? options) {
            ExcelImageExportOptions resolved = options?.Clone() ?? new ExcelImageExportOptions();
            resolved.ConditionalFormattingDate ??= System.DateTime.Today;
            OfficeImageExportOptions.ValidateScale(resolved.Scale, nameof(options));

            return resolved;
        }

        private static void WriteFile(string path, byte[] bytes) {
            if (string.IsNullOrWhiteSpace(path)) {
                throw new ArgumentException("Output path cannot be null or whitespace.", nameof(path));
            }

            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) {
                Directory.CreateDirectory(directory!);
            }

            OfficeFileCommit.WriteAllBytes(fullPath, bytes);
        }

        private static void WriteStream(Stream stream, byte[] bytes) {
            if (stream == null) {
                throw new ArgumentNullException(nameof(stream));
            }

            stream.Write(bytes, 0, bytes.Length);
        }
    }
}
