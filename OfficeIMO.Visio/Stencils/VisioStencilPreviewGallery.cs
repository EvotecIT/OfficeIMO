using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Options for exporting extracted stencil preview payloads as a browsable review gallery.
    /// </summary>
    public sealed class VisioStencilPreviewGalleryOptions {
        /// <summary>
        /// Gets or sets the gallery title. When null, a title is derived from the package file name.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the subdirectory that receives extracted preview payload files.
        /// </summary>
        public string PreviewDirectoryName { get; set; } = "previews";

        /// <summary>
        /// Gets or sets the generated HTML index file name.
        /// </summary>
        public string IndexFileName { get; set; } = "index.html";

        /// <summary>
        /// Gets or sets whether the HTML gallery index should be written.
        /// </summary>
        public bool WriteIndex { get; set; } = true;

        /// <summary>
        /// Gets or sets whether browser-renderable preview payloads should receive deterministic SVG thumbnail wrappers.
        /// </summary>
        public bool WriteBrowserRenderableThumbnails { get; set; } = true;

        /// <summary>
        /// Gets or sets the subdirectory that receives generated SVG thumbnail wrappers.
        /// </summary>
        public string ThumbnailDirectoryName { get; set; } = "thumbnails";

        /// <summary>
        /// Gets or sets generated thumbnail width in pixels.
        /// </summary>
        public int ThumbnailWidth { get; set; } = 220;

        /// <summary>
        /// Gets or sets generated thumbnail height in pixels.
        /// </summary>
        public int ThumbnailHeight { get; set; } = 160;
    }

    /// <summary>
    /// Result produced when exporting package stencil preview payloads for review.
    /// </summary>
    public sealed class VisioStencilPreviewGallery {
        internal VisioStencilPreviewGallery(string packagePath, string outputDirectory, string previewDirectory, string? thumbnailDirectory, string? indexPath, IReadOnlyList<VisioStencilPreviewGalleryEntry> entries) {
            PackagePath = packagePath;
            OutputDirectory = outputDirectory;
            PreviewDirectory = previewDirectory;
            ThumbnailDirectory = thumbnailDirectory;
            IndexPath = indexPath;
            Entries = entries;
        }

        /// <summary>Source Visio package path.</summary>
        public string PackagePath { get; }

        /// <summary>Gallery output directory.</summary>
        public string OutputDirectory { get; }

        /// <summary>Directory containing extracted preview payload files.</summary>
        public string PreviewDirectory { get; }

        /// <summary>Directory containing generated SVG thumbnail wrappers, when enabled.</summary>
        public string? ThumbnailDirectory { get; }

        /// <summary>Generated HTML index path, when written.</summary>
        public string? IndexPath { get; }

        /// <summary>Extracted preview entries.</summary>
        public IReadOnlyList<VisioStencilPreviewGalleryEntry> Entries { get; }

        /// <summary>Number of preview payloads that the generated gallery renders inline.</summary>
        public int BrowserRenderableCount => Entries.Count(entry => entry.IsBrowserRenderable);

        /// <summary>Number of generated thumbnail artifacts.</summary>
        public int ThumbnailCount => Entries.Count(entry => entry.HasThumbnail);
    }

    /// <summary>
    /// One extracted stencil preview payload in a review gallery.
    /// </summary>
    public sealed class VisioStencilPreviewGalleryEntry {
        internal VisioStencilPreviewGalleryEntry(VisioStencilPreviewImageData image, string filePath, string relativePath, string? thumbnailFilePath, string? thumbnailRelativePath) {
            Image = image;
            FilePath = filePath;
            RelativePath = relativePath;
            ThumbnailFilePath = thumbnailFilePath;
            ThumbnailRelativePath = thumbnailRelativePath;
        }

        /// <summary>Extracted preview payload and source master metadata.</summary>
        public VisioStencilPreviewImageData Image { get; }

        /// <summary>Saved preview payload path.</summary>
        public string FilePath { get; }

        /// <summary>Path from the gallery index to the saved preview payload.</summary>
        public string RelativePath { get; }

        /// <summary>Generated SVG thumbnail path, when available.</summary>
        public string? ThumbnailFilePath { get; }

        /// <summary>Path from the gallery index to the generated SVG thumbnail, when available.</summary>
        public string? ThumbnailRelativePath { get; }

        /// <summary>Whether this entry has a generated SVG thumbnail artifact.</summary>
        public bool HasThumbnail => !string.IsNullOrWhiteSpace(ThumbnailFilePath);

        /// <summary>Whether the generated gallery can safely render the payload inline.</summary>
        public bool IsBrowserRenderable => IsBrowserRenderableExtension(Image.PreviewImage.Extension);

        internal static bool IsBrowserRenderableExtension(string? extension) =>
            OfficeImageInfo.IsBrowserPreviewSafeExtension(extension);
    }

    internal static class VisioStencilPreviewGalleryWriter {
        internal static VisioStencilPreviewGallery Create(
            string packagePath,
            string outputDirectory,
            IReadOnlyList<VisioStencilPreviewImageData> images,
            VisioStencilPreviewGalleryOptions options) {
            string fullPackagePath = Path.GetFullPath(packagePath);
            string fullOutputDirectory = Path.GetFullPath(outputDirectory);
            string previewDirectory = Path.Combine(fullOutputDirectory, options.PreviewDirectoryName);
            string? thumbnailDirectory = options.WriteBrowserRenderableThumbnails
                ? Path.Combine(fullOutputDirectory, options.ThumbnailDirectoryName)
                : null;
            Directory.CreateDirectory(previewDirectory);
            if (thumbnailDirectory != null) {
                Directory.CreateDirectory(thumbnailDirectory);
            }

            List<VisioStencilPreviewGalleryEntry> entries = new();
            foreach (VisioStencilPreviewImageData image in images.OrderBy(image => image.MasterNameU, StringComparer.OrdinalIgnoreCase)) {
                string filePath = SaveGalleryPreviewPayload(previewDirectory, image);
                string relativePath = Path.Combine(options.PreviewDirectoryName, Path.GetFileName(filePath))
                    .Replace(Path.DirectorySeparatorChar, '/')
                    .Replace(Path.AltDirectorySeparatorChar, '/');
                string? thumbnailFilePath = null;
                string? thumbnailRelativePath = null;
                if (thumbnailDirectory != null &&
                    VisioStencilPreviewGalleryEntry.IsBrowserRenderableExtension(image.PreviewImage.Extension)) {
                    thumbnailFilePath = WriteThumbnail(thumbnailDirectory, image, options);
                    thumbnailRelativePath = Path.Combine(options.ThumbnailDirectoryName, Path.GetFileName(thumbnailFilePath))
                        .Replace(Path.DirectorySeparatorChar, '/')
                        .Replace(Path.AltDirectorySeparatorChar, '/');
                }

                entries.Add(new VisioStencilPreviewGalleryEntry(image, filePath, relativePath, thumbnailFilePath, thumbnailRelativePath));
            }

            string? indexPath = null;
            if (options.WriteIndex) {
                indexPath = Path.Combine(fullOutputDirectory, options.IndexFileName);
                WriteIndex(indexPath, fullPackagePath, entries, options);
            }

            return new VisioStencilPreviewGallery(fullPackagePath, fullOutputDirectory, previewDirectory, thumbnailDirectory, indexPath, entries.AsReadOnly());
        }

        private static string SaveGalleryPreviewPayload(string previewDirectory, VisioStencilPreviewImageData image) {
            string extension = image.PreviewImage.Extension?.TrimStart('.') ?? string.Empty;
            if (!string.Equals(extension, "svg", StringComparison.OrdinalIgnoreCase)) {
                return image.SaveToDirectory(previewDirectory);
            }

            // Preserve the source bytes for review, but use a text extension so a gallery host cannot serve an
            // attacker-controlled SVG as an executable top-level image document by extension alone.
            string path = Path.Combine(previewDirectory, image.SuggestedFileName + ".txt");
            image.Save(path);
            return path;
        }

        internal static void ValidateOptions(VisioStencilPreviewGalleryOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name cannot be null or whitespace.", nameof(options));
            if (Path.IsPathRooted(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name must be relative.", nameof(options));
            if (ContainsParentSegment(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name cannot contain parent directory segments.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.ThumbnailDirectoryName)) throw new ArgumentException("Thumbnail directory name cannot be null or whitespace.", nameof(options));
            if (Path.IsPathRooted(options.ThumbnailDirectoryName)) throw new ArgumentException("Thumbnail directory name must be relative.", nameof(options));
            if (ContainsParentSegment(options.ThumbnailDirectoryName)) throw new ArgumentException("Thumbnail directory name cannot contain parent directory segments.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.IndexFileName)) throw new ArgumentException("Index file name cannot be null or whitespace.", nameof(options));
            if (Path.IsPathRooted(options.IndexFileName)) throw new ArgumentException("Index file name must be relative.", nameof(options));
            if (ContainsParentSegment(options.IndexFileName)) throw new ArgumentException("Index file name cannot contain parent directory segments.", nameof(options));
            if (options.IndexFileName.IndexOf(Path.DirectorySeparatorChar) >= 0 || options.IndexFileName.IndexOf(Path.AltDirectorySeparatorChar) >= 0) throw new ArgumentException("Index file name cannot contain directory separators.", nameof(options));
            if (options.ThumbnailWidth <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Thumbnail width must be positive.");
            if (options.ThumbnailHeight <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Thumbnail height must be positive.");
        }

        private static bool ContainsParentSegment(string path) {
            return path
                .Split(new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)
                .Any(segment => segment == "..");
        }

        private static void WriteIndex(string indexPath, string packagePath, IReadOnlyList<VisioStencilPreviewGalleryEntry> entries, VisioStencilPreviewGalleryOptions options) {
            Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
            string title = string.IsNullOrWhiteSpace(options.Title)
                ? "Stencil Preview Gallery - " + Path.GetFileName(packagePath)
                : options.Title!;
            int renderable = entries.Count(entry => entry.IsBrowserRenderable);
            string contentTypes = string.Join(", ", entries
                .Select(entry => entry.Image.PreviewImage.ContentType)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase));

            StringBuilder builder = new();
            builder.AppendLine("<!doctype html>");
            builder.AppendLine("<html lang=\"en\">");
            builder.AppendLine("<head>");
            builder.AppendLine("  <meta charset=\"utf-8\">");
            builder.AppendLine("  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">");
            builder.AppendLine("  <title>" + Escape(title) + "</title>");
            builder.AppendLine("  <style>");
            builder.AppendLine("    :root { color-scheme: light; --ink: #1f3040; --muted: #657586; --line: #d3e0ec; --panel: #f8fbfe; --accent: #2563eb; }");
            builder.AppendLine("    * { box-sizing: border-box; }");
            builder.AppendLine("    body { margin: 0; font: 14px/1.45 Aptos, Segoe UI, Arial, sans-serif; color: var(--ink); background: #ffffff; }");
            builder.AppendLine("    header { padding: 32px 40px 20px; border-bottom: 1px solid var(--line); background: linear-gradient(180deg, #f8fbfe 0%, #ffffff 100%); }");
            builder.AppendLine("    h1 { margin: 0 0 8px; font-size: 28px; font-weight: 700; letter-spacing: 0; }");
            builder.AppendLine("    .meta { color: var(--muted); overflow-wrap: anywhere; }");
            builder.AppendLine("    .stats { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 18px; }");
            builder.AppendLine("    .stat { border: 1px solid var(--line); border-radius: 8px; padding: 8px 12px; background: #fff; }");
            builder.AppendLine("    main { padding: 28px 40px 40px; }");
            builder.AppendLine("    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; }");
            builder.AppendLine("    article { border: 1px solid var(--line); border-radius: 8px; background: var(--panel); overflow: hidden; }");
            builder.AppendLine("    .preview { height: 150px; display: grid; place-items: center; background: #fff; border-bottom: 1px solid var(--line); }");
            builder.AppendLine("    .preview img { max-width: 88%; max-height: 112px; object-fit: contain; }");
            builder.AppendLine("    .fallback { display: grid; place-items: center; width: 96px; height: 72px; border: 1px solid var(--line); border-radius: 8px; color: var(--accent); font-weight: 700; background: #f5f9ff; text-transform: uppercase; }");
            builder.AppendLine("    .body { padding: 14px 16px 16px; }");
            builder.AppendLine("    h2 { margin: 0 0 10px; font-size: 16px; }");
            builder.AppendLine("    dl { display: grid; grid-template-columns: 88px minmax(0, 1fr); gap: 4px 10px; margin: 0; }");
            builder.AppendLine("    dt { color: var(--muted); }");
            builder.AppendLine("    dd { margin: 0; overflow-wrap: anywhere; }");
            builder.AppendLine("    a { color: var(--accent); text-decoration: none; }");
            builder.AppendLine("    a:hover { text-decoration: underline; }");
            builder.AppendLine("  </style>");
            builder.AppendLine("</head>");
            builder.AppendLine("<body>");
            builder.AppendLine("  <header>");
            builder.AppendLine("    <h1>" + Escape(title) + "</h1>");
            builder.AppendLine("    <div class=\"meta\">" + Escape(packagePath) + "</div>");
            builder.AppendLine("    <div class=\"stats\">");
            builder.AppendLine("      <div class=\"stat\"><strong>" + entries.Count.ToString(CultureInfo.InvariantCulture) + "</strong> previews</div>");
            builder.AppendLine("      <div class=\"stat\"><strong>" + renderable.ToString(CultureInfo.InvariantCulture) + "</strong> browser-renderable</div>");
            builder.AppendLine("      <div class=\"stat\"><strong>" + entries.Count(entry => entry.HasThumbnail).ToString(CultureInfo.InvariantCulture) + "</strong> thumbnails</div>");
            builder.AppendLine("      <div class=\"stat\">" + Escape(string.IsNullOrWhiteSpace(contentTypes) ? "unknown content types" : contentTypes) + "</div>");
            builder.AppendLine("    </div>");
            builder.AppendLine("  </header>");
            builder.AppendLine("  <main>");
            builder.AppendLine("    <section class=\"grid\">");
            foreach (VisioStencilPreviewGalleryEntry entry in entries) {
                AppendEntry(builder, entry);
            }
            builder.AppendLine("    </section>");
            builder.AppendLine("  </main>");
            builder.AppendLine("</body>");
            builder.AppendLine("</html>");
            OfficeFileCommit.WriteAllBytes(indexPath, new UTF8Encoding(false).GetBytes(builder.ToString()));
        }

        private static void AppendEntry(StringBuilder builder, VisioStencilPreviewGalleryEntry entry) {
            string displayName = string.IsNullOrWhiteSpace(entry.Image.MasterName) ? entry.Image.MasterNameU : entry.Image.MasterName!;
            string extension = string.IsNullOrWhiteSpace(entry.Image.PreviewImage.Extension)
                ? "bin"
                : entry.Image.PreviewImage.Extension!.TrimStart('.');

            builder.AppendLine("      <article>");
            builder.AppendLine("        <div class=\"preview\">");
            if (!string.IsNullOrWhiteSpace(entry.ThumbnailRelativePath)) {
                builder.AppendLine("          <img src=\"" + Escape(entry.ThumbnailRelativePath!) + "\" alt=\"" + Escape(displayName) + "\">");
            } else if (entry.IsBrowserRenderable) {
                builder.AppendLine("          <img src=\"" + Escape(entry.RelativePath) + "\" alt=\"" + Escape(displayName) + "\">");
            } else {
                builder.AppendLine("          <div class=\"fallback\">" + Escape(extension) + "</div>");
            }

            builder.AppendLine("        </div>");
            builder.AppendLine("        <div class=\"body\">");
            builder.AppendLine("          <h2>" + Escape(displayName) + "</h2>");
            builder.AppendLine("          <dl>");
            AppendDefinition(builder, "NameU", entry.Image.MasterNameU);
            AppendDefinition(builder, "Master", entry.Image.MasterId);
            AppendDefinition(builder, "Type", entry.Image.PreviewImage.ContentType ?? string.Empty);
            AppendDefinition(builder, "Bytes", entry.Image.ByteLength.ToString(CultureInfo.InvariantCulture));
            AppendDefinition(builder, "Target", entry.Image.PreviewImage.Target);
            builder.AppendLine("            <dt>File</dt><dd><a download href=\"" + Escape(entry.RelativePath) + "\">" + Escape(Path.GetFileName(entry.FilePath)) + "</a></dd>");
            if (!string.IsNullOrWhiteSpace(entry.ThumbnailRelativePath)) {
                builder.AppendLine("            <dt>Thumb</dt><dd><a href=\"" + Escape(entry.ThumbnailRelativePath!) + "\">" + Escape(Path.GetFileName(entry.ThumbnailFilePath!)) + "</a></dd>");
            }
            builder.AppendLine("          </dl>");
            builder.AppendLine("        </div>");
            builder.AppendLine("      </article>");
        }

        private static string WriteThumbnail(string thumbnailDirectory, VisioStencilPreviewImageData image, VisioStencilPreviewGalleryOptions options) {
            string fileName = Path.GetFileNameWithoutExtension(image.SuggestedFileName) + ".thumbnail.svg";
            string path = Path.Combine(thumbnailDirectory, fileName);
            string displayName = string.IsNullOrWhiteSpace(image.MasterName) ? image.MasterNameU : image.MasterName!;
            string contentType = string.IsNullOrWhiteSpace(image.PreviewImage.ContentType)
                ? OfficeImageInfo.GetMimeTypeFromExtension(image.PreviewImage.Extension)
                : image.PreviewImage.ContentType!;
            string dataUri = OfficeSvgImageRenderer.CreateDataUri(contentType, image.ToBytes());
            string width = options.ThumbnailWidth.ToString(CultureInfo.InvariantCulture);
            string height = options.ThumbnailHeight.ToString(CultureInfo.InvariantCulture);
            double imageWidth = Math.Max(1, options.ThumbnailWidth - 28);
            double imageHeight = Math.Max(1, options.ThumbnailHeight - 42);

            StringBuilder builder = new();
            builder.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.AppendLine("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"" + width + "\" height=\"" + height + "\" viewBox=\"0 0 " + width + " " + height + "\" role=\"img\" aria-label=\"" + Escape(displayName) + "\">");
            builder.Append("  ").AppendRectElement(0D, 0D, options.ThumbnailWidth, options.ThumbnailHeight, 8D, 8D, " fill=\"#FFFFFF\"").AppendLine();
            builder.Append("  ").AppendRectElement(0.5D, 0.5D, options.ThumbnailWidth - 1D, options.ThumbnailHeight - 1D, 7.5D, 7.5D, " fill=\"none\" stroke=\"#D3E0EC\"").AppendLine();
            builder.Append("  ");
            OfficeSvgImageRenderer.AppendImageInViewport(
                builder,
                dataUri,
                new OfficeImageProjection(new OfficeImagePlacement(14D, 12D, imageWidth, imageHeight)),
                "visio-thumbnail-image-clip",
                new OfficeImagePlacement(14D, 12D, imageWidth, imageHeight),
                preserveAspectRatio: "xMidYMid meet").AppendLine();
            builder.Append("  ").AppendSvgTextElement(
                displayName,
                14D,
                options.ThumbnailHeight - 14D,
                12D,
                OfficeColor.FromRgb(101, 117, 134),
                "Aptos, Segoe UI, Arial, sans-serif",
                12D,
                OfficeTextAlignment.Left).AppendLine();
            builder.AppendLine("</svg>");
            OfficeFileCommit.WriteAllBytes(path, new UTF8Encoding(false).GetBytes(builder.ToString()));
            return path;
        }

        private static void AppendDefinition(StringBuilder builder, string name, string value) {
            builder.AppendLine("            <dt>" + Escape(name) + "</dt><dd>" + Escape(value) + "</dd>");
        }

        private static string Escape(string? value) => OfficeSvgFormatting.Escape(value);

    }
}
