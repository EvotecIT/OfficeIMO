using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Renders the requested page to dependency-free JPEG bytes.</summary>
        public byte[] ToJpeg(WordImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Jpeg, options).Bytes;

        /// <summary>Renders the requested page to dependency-free TIFF bytes.</summary>
        public byte[] ToTiff(WordImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Tiff, options).Bytes;

        /// <summary>Renders the requested page to dependency-free lossless WebP bytes.</summary>
        public byte[] ToWebp(WordImageExportOptions? options = null) =>
            ExportImage(OfficeImageExportFormat.Webp, options).Bytes;

        /// <summary>Saves the requested page as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsJpeg().Save(path);

        /// <summary>Saves the requested page as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsTiff().Save(path);

        /// <summary>Saves the requested page as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(string path, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsWebp().Save(path);

        /// <summary>Writes the requested page as JPEG.</summary>
        public OfficeImageExportResult SaveAsJpeg(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsJpeg().Save(stream);

        /// <summary>Writes the requested page as TIFF.</summary>
        public OfficeImageExportResult SaveAsTiff(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsTiff().Save(stream);

        /// <summary>Writes the requested page as lossless WebP.</summary>
        public OfficeImageExportResult SaveAsWebp(Stream stream, WordImageExportOptions? options = null) =>
            new WordDocumentImageExportBuilder(this, options).AsWebp().Save(stream);

        /// <summary>Asynchronously saves the requested page as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsJpeg().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the requested page as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsTiff().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously saves the requested page as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(string path, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsWebp().SaveAsync(path, cancellationToken);

        /// <summary>Asynchronously writes the requested page as JPEG.</summary>
        public Task<OfficeImageExportResult> SaveAsJpegAsync(Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsJpeg().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the requested page as TIFF.</summary>
        public Task<OfficeImageExportResult> SaveAsTiffAsync(Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsTiff().SaveAsync(stream, cancellationToken);

        /// <summary>Asynchronously writes the requested page as lossless WebP.</summary>
        public Task<OfficeImageExportResult> SaveAsWebpAsync(Stream stream, WordImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
            new WordDocumentImageExportBuilder(this, options).AsWebp().SaveAsync(stream, cancellationToken);
    }
}
