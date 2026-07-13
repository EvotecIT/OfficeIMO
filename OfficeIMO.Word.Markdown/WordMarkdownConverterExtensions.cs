using OfficeIMO.Markdown;
using OfficeIMO.Drawing.Internal;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Extension methods enabling Markdown conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static class WordMarkdownConverterExtensions {
        /// <summary>
        /// Synchronously saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        public static WordToMarkdownResult SaveAsMarkdown(this WordDocument document, string path, WordToMarkdownOptions? options = null) {
            var effectiveOptions = CreateFileSaveOptions(path, options);
            WordToMarkdownResult result = ConvertToMarkdownResult(document, effectiveOptions);
            OfficeFileCommit.WriteAllBytes(path, Encoding.UTF8.GetBytes(RenderMarkdown(result.Value)));
            return result;
        }

        /// <summary>
        /// Synchronously saves the document as Markdown to the provided stream.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        public static WordToMarkdownResult SaveAsMarkdown(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null) {
            WordToMarkdownResult result = ConvertToMarkdownResult(document, CopyOptions(options ?? new WordToMarkdownOptions()));
            OfficeStreamWriter.WriteAllBytes(stream, Encoding.UTF8.GetBytes(RenderMarkdown(result.Value)));
            return result;
        }

        /// <summary>
        /// Asynchronously saves the document as a Markdown file at the specified path.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="path">Destination file path.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A task representing the asynchronous save operation.</returns>
        public static async Task<WordToMarkdownResult> SaveAsMarkdownAsync(this WordDocument document, string path, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            var effectiveOptions = CreateFileSaveOptions(path, options);
            cancellationToken.ThrowIfCancellationRequested();
            WordToMarkdownResult result = ConvertToMarkdownResult(document, effectiveOptions, cancellationToken);
            await OfficeFileCommit.WriteAllBytesAsync(
                path,
                Encoding.UTF8.GetBytes(RenderMarkdown(result.Value)),
                cancellationToken: cancellationToken).ConfigureAwait(false);
            return result;
        }

        /// <summary>
        /// Asynchronously saves the document as Markdown to the provided stream.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> to convert to Markdown.</param>
        /// <param name="stream">Target stream.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor for cancellation requests.</param>
        /// <returns>A task representing the asynchronous save operation.</returns>
        public static async Task<WordToMarkdownResult> SaveAsMarkdownAsync(this WordDocument document, Stream stream, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            WordToMarkdownResult result = ConvertToMarkdownResult(
                document,
                CopyOptions(options ?? new WordToMarkdownOptions()),
                cancellationToken);
            await OfficeStreamWriter.WriteAllBytesAsync(
                stream,
                Encoding.UTF8.GetBytes(RenderMarkdown(result.Value)),
                cancellationToken).ConfigureAwait(false);
            return result;
        }

        private static WordToMarkdownOptions CreateFileSaveOptions(string markdownPath, WordToMarkdownOptions? options) {
            var effectiveOptions = CopyOptions(options ?? new WordToMarkdownOptions());
            if (effectiveOptions.ImageExportMode == ImageExportMode.File && string.IsNullOrEmpty(effectiveOptions.ImageDirectory)) {
                effectiveOptions.ImageDirectory = Path.GetDirectoryName(markdownPath) ?? Directory.GetCurrentDirectory();
            }

            if (effectiveOptions.VisualFallbackMode != MarkdownVisualFallbackMode.SvgFile) {
                return effectiveOptions;
            }

            string markdownDirectory = Path.GetDirectoryName(markdownPath) ?? Directory.GetCurrentDirectory();
            if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackDirectory)) {
                string markdownFileName = Path.GetFileNameWithoutExtension(markdownPath);
                string sidecarDirectoryName = string.IsNullOrEmpty(markdownFileName) ? "assets" : markdownFileName + ".assets";
                effectiveOptions.VisualFallbackDirectory = Path.Combine(markdownDirectory, sidecarDirectoryName);
                if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackPathPrefix)) {
                    effectiveOptions.VisualFallbackPathPrefix = sidecarDirectoryName;
                }
                return effectiveOptions;
            }

            if (string.IsNullOrEmpty(effectiveOptions.VisualFallbackPathPrefix)) {
                effectiveOptions.VisualFallbackPathPrefix = MakeMarkdownRelativePath(markdownDirectory, effectiveOptions.VisualFallbackDirectory!);
            }

            return effectiveOptions;
        }

        private static WordToMarkdownOptions CopyOptions(WordToMarkdownOptions source) {
            return new WordToMarkdownOptions {
                FontFamily = source.FontFamily,
                EnableUnderline = source.EnableUnderline,
                EnableHighlight = source.EnableHighlight,
                ImageExportMode = source.ImageExportMode,
                ImageDirectory = source.ImageDirectory,
                FallbackExternalImagesToLinks = source.FallbackExternalImagesToLinks,
                PageBreakMode = source.PageBreakMode,
                UnsupportedContentMode = source.UnsupportedContentMode,
                VisualFallbackMode = source.VisualFallbackMode,
                VisualFallbackDirectory = source.VisualFallbackDirectory,
                VisualFallbackPathPrefix = source.VisualFallbackPathPrefix,
                OnWarning = source.OnWarning,
                IncludeHeadersAndFootersAsSemanticBlocks = source.IncludeHeadersAndFootersAsSemanticBlocks
            };
        }

        private static WordToMarkdownResult ConvertToMarkdownResult(
            WordDocument document,
            WordToMarkdownOptions effectiveOptions,
            CancellationToken cancellationToken = default) {
            var diagnostics = new List<WordMarkdownConversionDiagnostic>();
            Action<string>? callerWarning = effectiveOptions.OnWarning;
            effectiveOptions.OnWarning = message => {
                diagnostics.Add(new WordMarkdownConversionDiagnostic(
                    "WordToMarkdownWarning",
                    message,
                    WordMarkdownConversionLossKind.Approximation));
                callerWarning?.Invoke(message);
            };

            MarkdownDoc value = new WordToMarkdownConverter().ConvertToDocument(document, effectiveOptions, cancellationToken);
            return new WordToMarkdownResult(value, new WordMarkdownConversionReport(diagnostics));
        }

        private static string RenderMarkdown(MarkdownDoc document) {
            string markdown = document.ToMarkdown();
            return string.IsNullOrEmpty(markdown)
                ? string.Empty
                : markdown.Replace("\r\n", "\n").TrimEnd('\n');
        }

        private static string MakeMarkdownRelativePath(string baseDirectory, string targetDirectory) {
            try {
                string baseFullPath = Path.GetFullPath(baseDirectory);
                if (!baseFullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) &&
                    !baseFullPath.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
                    baseFullPath += Path.DirectorySeparatorChar;
                }

                string targetFullPath = Path.GetFullPath(targetDirectory);
                var baseUri = new Uri(baseFullPath);
                var targetUri = new Uri(targetFullPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                    ? targetFullPath
                    : targetFullPath + Path.DirectorySeparatorChar);
                string relative = Uri.UnescapeDataString(baseUri.MakeRelativeUri(targetUri).ToString());
                return relative.TrimEnd('/').Replace('\\', '/');
            } catch {
                return Path.GetFileName(targetDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            }
        }

        /// <summary>
        /// Converts the document to a Markdown string.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor during conversion.</param>
        /// <returns>Markdown representation of the document.</returns>
        public static string ToMarkdown(this WordDocument document, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            return RenderMarkdown(document.ToMarkdownDocumentResult(options, cancellationToken).Value);
        }

        /// <summary>
        /// Converts the Word document to a typed Markdown document with a structured fidelity report.
        /// </summary>
        public static WordToMarkdownResult ToMarkdownDocumentResult(
            this WordDocument document,
            WordToMarkdownOptions? options = null,
            CancellationToken cancellationToken = default) {
            return ConvertToMarkdownResult(document, CopyOptions(options ?? new WordToMarkdownOptions()), cancellationToken);
        }

        /// <summary>
        /// Converts the document into a typed <see cref="MarkdownDoc"/> without flattening through markdown text first.
        /// </summary>
        /// <param name="document">Document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor during conversion.</param>
        /// <returns>Typed markdown document.</returns>
        public static MarkdownDoc ToMarkdownDocument(this WordDocument document, WordToMarkdownOptions? options = null, CancellationToken cancellationToken = default) {
            return document.ToMarkdownDocumentResult(options, cancellationToken).Value;
        }

        /// <summary>
        /// Converts a typed Markdown document to Word with a structured fidelity report.
        /// </summary>
        public static MarkdownToWordResult ToWordDocumentResult(
            this MarkdownDoc markdown,
            MarkdownToWordOptions? options = null,
            CancellationToken cancellationToken = default) {
            MarkdownToWordOptions effectiveOptions = CopyOptions(options ?? new MarkdownToWordOptions());
            var diagnostics = new List<WordMarkdownConversionDiagnostic>();
            Action<string>? callerWarning = effectiveOptions.OnWarning;
            effectiveOptions.OnWarning = message => {
                diagnostics.Add(new WordMarkdownConversionDiagnostic(
                    "MarkdownToWordWarning",
                    message,
                    WordMarkdownConversionLossKind.Approximation));
                callerWarning?.Invoke(message);
            };

            WordDocument value = new MarkdownToWordConverter().Convert(markdown, effectiveOptions, cancellationToken);
            return new MarkdownToWordResult(value, new WordMarkdownConversionReport(diagnostics));
        }

        /// <summary>
        /// Creates a new Word document directly from a typed <see cref="MarkdownDoc"/>.
        /// </summary>
        /// <param name="markdown">Typed markdown document to convert.</param>
        /// <param name="options">Optional conversion options.</param>
        /// <param name="cancellationToken">Token to monitor during conversion.</param>
        /// <returns>A new <see cref="WordDocument"/> instance.</returns>
        public static WordDocument ToWordDocument(this MarkdownDoc markdown, MarkdownToWordOptions? options = null, CancellationToken cancellationToken = default) {
            return markdown.ToWordDocumentResult(options, cancellationToken).Value;
        }

        private static MarkdownToWordOptions CopyOptions(MarkdownToWordOptions source) {
            var copy = new MarkdownToWordOptions {
                Theme = source.Theme,
                ApplyDefaultTheme = source.ApplyDefaultTheme,
                FontFamily = source.FontFamily,
                BaseUri = source.BaseUri,
                OnWarning = source.OnWarning,
                DefaultPageSize = source.DefaultPageSize,
                DefaultOrientation = source.DefaultOrientation,
                AllowLocalImages = source.AllowLocalImages,
                RemoteImageResolver = source.RemoteImageResolver,
                AllowDataUriImages = source.AllowDataUriImages,
                MaxDataUriImageBytes = source.MaxDataUriImageBytes,
                MaximumRemoteImageBytes = source.MaximumRemoteImageBytes,
                ImageUrlValidator = source.ImageUrlValidator,
                FallbackRemoteImagesToHyperlinks = source.FallbackRemoteImagesToHyperlinks,
                OnImageLayoutDiagnostic = source.OnImageLayoutDiagnostic,
                PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
                ReaderOptions = source.ReaderOptions?.Clone(),
                RenderFrontMatter = source.RenderFrontMatter
            };

            foreach (string directory in source.AllowedImageDirectories) copy.AllowedImageDirectories.Add(directory);
            copy.AllowedImageSchemes.Clear();
            foreach (string scheme in source.AllowedImageSchemes) copy.AllowedImageSchemes.Add(scheme);
            copy.ImageLayout.FitMode = source.ImageLayout.FitMode;
            copy.ImageLayout.HintPrecedence = source.ImageLayout.HintPrecedence;
            copy.ImageLayout.MaxWidthPixels = source.ImageLayout.MaxWidthPixels;
            copy.ImageLayout.MaxHeightPixels = source.ImageLayout.MaxHeightPixels;
            copy.ImageLayout.MaxWidthPercentOfContent = source.ImageLayout.MaxWidthPercentOfContent;
            copy.ImageLayout.AllowUpscale = source.ImageLayout.AllowUpscale;
            return copy;
        }

        /// <summary>
        /// Inserts a typed Markdown document into an existing Word template document.
        /// </summary>
        /// <param name="markdown">Typed Markdown document to insert.</param>
        /// <param name="templateDocument">Template document to populate.</param>
        /// <param name="options">Template insertion options.</param>
        /// <returns>The populated template document.</returns>
        public static WordDocument ToWordDocument(this MarkdownDoc markdown, WordDocument templateDocument, MarkdownToWordTemplateOptions? options = null) {
            var converter = new MarkdownToWordConverter();
            return converter.ConvertIntoTemplate(markdown, templateDocument, options ?? new MarkdownToWordTemplateOptions());
        }

    }
}
