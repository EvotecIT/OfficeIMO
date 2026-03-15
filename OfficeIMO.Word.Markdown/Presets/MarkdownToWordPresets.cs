using System.Collections.Generic;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

/// <summary>
/// Convenience factory methods for common markdown-to-Word host contracts.
/// </summary>
public static class MarkdownToWordPresets {
    private const int MinIntelligenceXVisualMaxWidthPx = 320;
    private const int MaxIntelligenceXVisualMaxWidthPx = 2000;
    private const int DefaultIntelligenceXVisualMaxWidthPx = 760;

    /// <summary>
    /// Creates the explicit IntelligenceX transcript markdown-to-Word conversion preset.
    /// </summary>
    public static MarkdownToWordOptions CreateIntelligenceXTranscript(
        IReadOnlyList<string>? allowedImageDirectories = null,
        int? visualMaxWidthPx = null) {
        var readerOptions = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);

        var options = new MarkdownToWordOptions {
            FontFamily = "Calibri",
            AllowLocalImages = allowedImageDirectories is { Count: > 0 },
            PreferNarrativeSingleLineDefinitions = true,
            ReaderOptions = readerOptions,
            MaxImageWidthPercentOfContent = 100d,
            FitImagesToPageContentWidth = true,
            MaxImageWidthPixels = NormalizeIntelligenceXVisualMaxWidthPx(visualMaxWidthPx)
        };

        if (allowedImageDirectories is { Count: > 0 }) {
            for (var i = 0; i < allowedImageDirectories.Count; i++) {
                var directory = allowedImageDirectories[i];
                if (string.IsNullOrWhiteSpace(directory)) {
                    continue;
                }

                if (!options.AllowedImageDirectories.Contains(directory)) {
                    options.AllowedImageDirectories.Add(directory);
                }
            }
        }

        return options;
    }

    private static int NormalizeIntelligenceXVisualMaxWidthPx(int? value) {
        if (!value.HasValue) {
            return DefaultIntelligenceXVisualMaxWidthPx;
        }

        var normalized = value.Value;
        if (normalized < MinIntelligenceXVisualMaxWidthPx) {
            return MinIntelligenceXVisualMaxWidthPx;
        }

        if (normalized > MaxIntelligenceXVisualMaxWidthPx) {
            return MaxIntelligenceXVisualMaxWidthPx;
        }

        return normalized;
    }
}
