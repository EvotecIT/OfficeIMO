using OfficeIMO.Markdown;
using OfficeIMO.Core.Internal;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Extension methods enabling Markdown conversions for <see cref="WordDocument"/> instances.
    /// </summary>
    public static partial class WordMarkdownConverterExtensions {








        private static WordMarkdownExportOptions CreateFileSaveOptions(string markdownPath, WordMarkdownExportOptions? options) {
            var effectiveOptions = CopyOptions(options ?? new WordMarkdownExportOptions());
            if (options == null) {
                // A saved Markdown file has a folder next to it, so images are written as files there instead of
                // inflating the text with base64 data URIs. In-memory conversion keeps data URIs.
                effectiveOptions.ImageExportMode = WordMarkdownImageExportMode.File;
            }
            if (effectiveOptions.ImageExportMode == WordMarkdownImageExportMode.File && string.IsNullOrEmpty(effectiveOptions.ImageDirectory)) {
                effectiveOptions.ImageDirectory = Path.GetDirectoryName(markdownPath) ?? Directory.GetCurrentDirectory();
            }

            if (effectiveOptions.VisualFallbackMode != WordMarkdownVisualFallbackMode.SvgFile) {
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

        internal static WordMarkdownExportOptions CopyOptions(WordMarkdownExportOptions source) {
            return new WordMarkdownExportOptions {
                FailOnLoss = source.FailOnLoss,
                DiagnosticSink = source.DiagnosticSink,
                FontFamily = source.FontFamily,
                EnableUnderline = source.EnableUnderline,
                EnableHighlight = source.EnableHighlight,
                ImageExportMode = source.ImageExportMode,
                ImageDirectory = source.ImageDirectory,
                MaxEmbeddedImageBytes = source.MaxEmbeddedImageBytes,
                MaxListNestingDepth = source.MaxListNestingDepth,
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
            WordMarkdownExportOptions effectiveOptions,
            CancellationToken cancellationToken = default) {
            var diagnostics = new List<WordMarkdownConversionDiagnostic>();
            var typedMessages = new HashSet<string>(StringComparer.Ordinal);
            Action<WordMarkdownConversionDiagnostic>? callerDiagnostic = effectiveOptions.DiagnosticSink;
            effectiveOptions.DiagnosticSink = diagnostic => {
                diagnostics.Add(diagnostic); typedMessages.Add(diagnostic.Message); callerDiagnostic?.Invoke(diagnostic);
            };
            Action<string>? callerWarning = effectiveOptions.OnWarning;
            effectiveOptions.OnWarning = message => {
                if (!typedMessages.Contains(message))
                    diagnostics.Add(new WordMarkdownConversionDiagnostic(
                        OfficeDiagnosticCodes.WordToMarkdownWarning, message, OfficeConversionLossKind.Approximation));
                callerWarning?.Invoke(message);
            };

            MarkdownDocument value = new WordToMarkdownConverter().ConvertToDocument(document, effectiveOptions, cancellationToken);
            return new WordToMarkdownResult(value, new WordMarkdownConversionReport(diagnostics));
        }

        private static string RenderMarkdown(MarkdownDocument document) {
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
        /// Converts the Word document to a typed Markdown document with a structured fidelity report.
        /// </summary>
        internal static WordToMarkdownResult ConvertToMarkdownWithEvidence(
            this WordDocument document,
            WordMarkdownExportOptions? options = null,
            CancellationToken cancellationToken = default) {
            return ConvertToMarkdownResult(document, CopyOptions(options ?? new WordMarkdownExportOptions()), cancellationToken);
        }



        /// <summary>
        /// Converts a typed Markdown document to Word with a structured fidelity report.
        /// </summary>
        internal static MarkdownToWordResult ConvertToWordWithEvidence(
            this MarkdownDocument markdown,
            WordMarkdownImportOptions? options = null,
            CancellationToken cancellationToken = default) {
            WordMarkdownImportOptions effectiveOptions = CopyOptions(options ?? new WordMarkdownImportOptions());
            var diagnostics = new List<WordMarkdownConversionDiagnostic>();
            Action<WordMarkdownConversionDiagnostic>? callerDiagnostic = effectiveOptions.DiagnosticSink;
            effectiveOptions.DiagnosticSink = diagnostic => {
                diagnostics.Add(diagnostic);
                callerDiagnostic?.Invoke(diagnostic);
            };
            Action<string>? callerWarning = effectiveOptions.OnWarning;
            effectiveOptions.OnWarning = message => {
                diagnostics.Add(new WordMarkdownConversionDiagnostic(
                    "markdown-to-word-warning",
                    message,
                    OfficeConversionLossKind.Approximation));
                callerWarning?.Invoke(message);
            };

            WordDocument value = effectiveOptions.Template == null
                ? new MarkdownToWordConverter().Convert(markdown, effectiveOptions, cancellationToken)
                : new MarkdownToWordConverter().ConvertIntoTemplate(markdown, effectiveOptions.Template, effectiveOptions);
            return new MarkdownToWordResult(value, new WordMarkdownConversionReport(diagnostics));
        }



        internal static WordMarkdownImportOptions CopyOptions(WordMarkdownImportOptions source) {
            var copy = new WordMarkdownImportOptions {
                Template = source.Template,
                ContentControlTag = source.ContentControlTag,
                ContentControlAlias = source.ContentControlAlias,
                BookmarkName = source.BookmarkName,
                ReplacePlaceholder = source.ReplacePlaceholder,
                RawHtmlRenderer = source.RawHtmlRenderer,
                RawHtmlParser = source.RawHtmlParser,
                DiagnosticSink = source.DiagnosticSink,
                Theme = source.Theme,
                ApplyDefaultTheme = source.ApplyDefaultTheme,
                FailOnLoss = source.FailOnLoss,
                FontFamily = source.FontFamily,
                BaseUri = source.BaseUri,
                OnWarning = source.OnWarning,
                MaxNestingDepth = source.MaxNestingDepth,
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



    }
}
