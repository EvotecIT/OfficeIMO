using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

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
    }

    /// <summary>
    /// Result produced when exporting package stencil preview payloads for review.
    /// </summary>
    public sealed class VisioStencilPreviewGallery {
        internal VisioStencilPreviewGallery(string packagePath, string outputDirectory, string previewDirectory, string? indexPath, IReadOnlyList<VisioStencilPreviewGalleryEntry> entries) {
            PackagePath = packagePath;
            OutputDirectory = outputDirectory;
            PreviewDirectory = previewDirectory;
            IndexPath = indexPath;
            Entries = entries;
        }

        /// <summary>Source Visio package path.</summary>
        public string PackagePath { get; }

        /// <summary>Gallery output directory.</summary>
        public string OutputDirectory { get; }

        /// <summary>Directory containing extracted preview payload files.</summary>
        public string PreviewDirectory { get; }

        /// <summary>Generated HTML index path, when written.</summary>
        public string? IndexPath { get; }

        /// <summary>Extracted preview entries.</summary>
        public IReadOnlyList<VisioStencilPreviewGalleryEntry> Entries { get; }

        /// <summary>Number of preview payloads that common browsers can render directly.</summary>
        public int BrowserRenderableCount => Entries.Count(entry => entry.IsBrowserRenderable);
    }

    /// <summary>
    /// One extracted stencil preview payload in a review gallery.
    /// </summary>
    public sealed class VisioStencilPreviewGalleryEntry {
        internal VisioStencilPreviewGalleryEntry(VisioStencilPreviewImageData image, string filePath, string relativePath) {
            Image = image;
            FilePath = filePath;
            RelativePath = relativePath;
        }

        /// <summary>Extracted preview payload and source master metadata.</summary>
        public VisioStencilPreviewImageData Image { get; }

        /// <summary>Saved preview payload path.</summary>
        public string FilePath { get; }

        /// <summary>Path from the gallery index to the saved preview payload.</summary>
        public string RelativePath { get; }

        /// <summary>Whether the payload extension is usually directly renderable in a browser.</summary>
        public bool IsBrowserRenderable => IsBrowserRenderableExtension(Image.PreviewImage.Extension);

        private static bool IsBrowserRenderableExtension(string? extension) {
            if (string.IsNullOrWhiteSpace(extension)) {
                return false;
            }

            return extension.TrimStart('.').ToLowerInvariant() switch {
                "png" or "jpg" or "jpeg" or "gif" or "svg" or "bmp" or "webp" => true,
                _ => false
            };
        }
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
            Directory.CreateDirectory(previewDirectory);

            List<VisioStencilPreviewGalleryEntry> entries = new();
            foreach (VisioStencilPreviewImageData image in images.OrderBy(image => image.MasterNameU, StringComparer.OrdinalIgnoreCase)) {
                string filePath = image.SaveToDirectory(previewDirectory);
                string relativePath = Path.Combine(options.PreviewDirectoryName, Path.GetFileName(filePath))
                    .Replace(Path.DirectorySeparatorChar, '/')
                    .Replace(Path.AltDirectorySeparatorChar, '/');
                entries.Add(new VisioStencilPreviewGalleryEntry(image, filePath, relativePath));
            }

            string? indexPath = null;
            if (options.WriteIndex) {
                indexPath = Path.Combine(fullOutputDirectory, options.IndexFileName);
                WriteIndex(indexPath, fullPackagePath, entries, options);
            }

            return new VisioStencilPreviewGallery(fullPackagePath, fullOutputDirectory, previewDirectory, indexPath, entries.AsReadOnly());
        }

        internal static void ValidateOptions(VisioStencilPreviewGalleryOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name cannot be null or whitespace.", nameof(options));
            if (Path.IsPathRooted(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name must be relative.", nameof(options));
            if (ContainsParentSegment(options.PreviewDirectoryName)) throw new ArgumentException("Preview directory name cannot contain parent directory segments.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.IndexFileName)) throw new ArgumentException("Index file name cannot be null or whitespace.", nameof(options));
            if (Path.IsPathRooted(options.IndexFileName)) throw new ArgumentException("Index file name must be relative.", nameof(options));
            if (ContainsParentSegment(options.IndexFileName)) throw new ArgumentException("Index file name cannot contain parent directory segments.", nameof(options));
            if (options.IndexFileName.Contains(Path.DirectorySeparatorChar) || options.IndexFileName.Contains(Path.AltDirectorySeparatorChar)) throw new ArgumentException("Index file name cannot contain directory separators.", nameof(options));
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
            File.WriteAllText(indexPath, builder.ToString(), new UTF8Encoding(false));
        }

        private static void AppendEntry(StringBuilder builder, VisioStencilPreviewGalleryEntry entry) {
            string displayName = string.IsNullOrWhiteSpace(entry.Image.MasterName) ? entry.Image.MasterNameU : entry.Image.MasterName!;
            string extension = string.IsNullOrWhiteSpace(entry.Image.PreviewImage.Extension)
                ? "bin"
                : entry.Image.PreviewImage.Extension!.TrimStart('.');

            builder.AppendLine("      <article>");
            builder.AppendLine("        <div class=\"preview\">");
            if (entry.IsBrowserRenderable) {
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
            builder.AppendLine("            <dt>File</dt><dd><a href=\"" + Escape(entry.RelativePath) + "\">" + Escape(Path.GetFileName(entry.FilePath)) + "</a></dd>");
            builder.AppendLine("          </dl>");
            builder.AppendLine("        </div>");
            builder.AppendLine("      </article>");
        }

        private static void AppendDefinition(StringBuilder builder, string name, string value) {
            builder.AppendLine("            <dt>" + Escape(name) + "</dt><dd>" + Escape(value) + "</dd>");
        }

        private static string Escape(string value) {
            return WebUtility.HtmlEncode(value);
        }
    }
}
