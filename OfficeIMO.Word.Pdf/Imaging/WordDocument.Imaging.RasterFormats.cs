using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public static partial class WordImageExportExtensions {
        /// <summary>Renders the requested page to dependency-free JPEG bytes.</summary>
        public static byte[] ToJpeg(this WordDocument document, WordImageExportOptions? options = null) =>
            document.ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders the requested page to dependency-free TIFF bytes.</summary>
        public static byte[] ToTiff(this WordDocument document, WordImageExportOptions? options = null) =>
            document.ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders the requested page to dependency-free lossless WebP bytes.</summary>
        public static byte[] ToWebp(this WordDocument document, WordImageExportOptions? options = null) =>
            document.ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves the requested page as JPEG.</summary>
        public static OfficeImageExportResult SaveAsJpeg(this WordDocument document, string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsJpeg().Save(path);

        /// <summary>Saves the requested page as TIFF.</summary>
        public static OfficeImageExportResult SaveAsTiff(this WordDocument document, string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsTiff().Save(path);

        /// <summary>Saves the requested page as lossless WebP.</summary>
        public static OfficeImageExportResult SaveAsWebp(this WordDocument document, string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsWebp().Save(path);

        /// <summary>Writes the requested page as JPEG.</summary>
        public static OfficeImageExportResult SaveAsJpeg(this WordDocument document, Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsJpeg().Save(stream);

        /// <summary>Writes the requested page as TIFF.</summary>
        public static OfficeImageExportResult SaveAsTiff(this WordDocument document, Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsTiff().Save(stream);

        /// <summary>Writes the requested page as lossless WebP.</summary>
        public static OfficeImageExportResult SaveAsWebp(this WordDocument document, Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(document, options).AsWebp().Save(stream);

        /// <summary>Asynchronously saves the requested page as JPEG.</summary>
        public static Task<OfficeImageExportResult> SaveAsJpegAsync(this WordDocument document, string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsJpeg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the requested page as TIFF.</summary>
        public static Task<OfficeImageExportResult> SaveAsTiffAsync(this WordDocument document, string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsTiff().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the requested page as lossless WebP.</summary>
        public static Task<OfficeImageExportResult> SaveAsWebpAsync(this WordDocument document, string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsWebp().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes the requested page as JPEG.</summary>
        public static Task<OfficeImageExportResult> SaveAsJpegAsync(this WordDocument document, Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsJpeg().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the requested page as TIFF.</summary>
        public static Task<OfficeImageExportResult> SaveAsTiffAsync(this WordDocument document, Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsTiff().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the requested page as lossless WebP.</summary>
        public static Task<OfficeImageExportResult> SaveAsWebpAsync(this WordDocument document, Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(document, options).AsWebp().SaveAsync(stream, cancellationToken);
    }
}
