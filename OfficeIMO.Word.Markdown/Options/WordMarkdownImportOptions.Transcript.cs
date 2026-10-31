using System.Collections.Generic;
using OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

public partial class WordMarkdownImportOptions {
    private const int MinIntelligenceXVisualMaxWidthPx = 320;
    private const int MaxIntelligenceXVisualMaxWidthPx = 2000;
    private const int DefaultIntelligenceXVisualMaxWidthPx = 760;

    /// <summary>
    /// Creates the explicit IntelligenceX transcript Markdown-to-Word conversion options.
    /// </summary>
    /// <param name="allowedImageDirectories">Local image roots permitted by the transcript import policy.</param>
    /// <param name="visualMaxWidth">Maximum visual width at 96 DPI, rounded to pixels and clamped to 320–2,000 pixels. Defaults to 760 pixels.</param>
    public static WordMarkdownImportOptions CreateIntelligenceXTranscript(
        IReadOnlyList<string>? allowedImageDirectories = null,
        OfficeLength? visualMaxWidth = null) {
        var readerOptions = MarkdownTranscriptPreparation.CreateIntelligenceXTranscriptReaderOptions(
            preservesGroupedDefinitionLikeParagraphs: false,
            visualFenceLanguageMode: MarkdownVisualFenceLanguageMode.IntelligenceXAliasFence);

        var options = new WordMarkdownImportOptions {
            FontFamily = "Calibri",
            AllowLocalImages = allowedImageDirectories is { Count: > 0 },
            PreferNarrativeSingleLineDefinitions = true,
            ReaderOptions = readerOptions,
            MaxImageWidthPercentOfContent = 100d,
            FitImagesToPageContentWidth = true,
            MaxImageWidth = global::OfficeIMO.OfficeLength.Pixels(NormalizeIntelligenceXVisualMaxWidth(visualMaxWidth))
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

    private static int NormalizeIntelligenceXVisualMaxWidth(OfficeLength? value) {
        double pixels = value?.ToPixels() ?? DefaultIntelligenceXVisualMaxWidthPx;
        return (int)System.Math.Round(System.Math.Max(MinIntelligenceXVisualMaxWidthPx,
            System.Math.Min(MaxIntelligenceXVisualMaxWidthPx, pixels)), System.MidpointRounding.AwayFromZero);
    }
}
