using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private static void CopyPowerPointTableCellToOdp(
        PowerPointTableCell source,
        OdpTableCell target,
        PowerPointOpenDocumentConversionOptions options,
        ref int paragraphs,
        ref int textRuns,
        ref int listParagraphs,
        ref int skippedBasicFormatting,
        ref int unsupportedHyperlinkTooltips,
        ref int unsupportedRunInteractions) {
        OdpParagraph? firstParagraph = target.Paragraphs.FirstOrDefault();
        bool useFirstParagraph = firstParagraph != null;

        OdpParagraph AddParagraph() {
            if (!useFirstParagraph) return target.AddParagraph();
            useFirstParagraph = false;
            firstParagraph!.Text = string.Empty;
            return firstParagraph;
        }

        CopyPowerPointParagraphsToOdp(source.Paragraphs, AddParagraph, options,
            ref paragraphs, ref textRuns, ref listParagraphs, ref skippedBasicFormatting,
            ref unsupportedHyperlinkTooltips, ref unsupportedRunInteractions);
    }

    private static void CopyPowerPointParagraphsToOdp(
        IReadOnlyList<PowerPointParagraph> sourceParagraphs,
        Func<OdpParagraph> addParagraph,
        PowerPointOpenDocumentConversionOptions options,
        ref int paragraphs,
        ref int textRuns,
        ref int listParagraphs,
        ref int skippedBasicFormatting,
        ref int unsupportedHyperlinkTooltips,
        ref int unsupportedRunInteractions) {
        foreach (PowerPointParagraph sourceParagraph in sourceParagraphs) {
            OdpParagraph targetParagraph = addParagraph();
            ApplyPowerPointParagraphLayout(sourceParagraph, targetParagraph);
            IReadOnlyList<PowerPointTextRun> runs = sourceParagraph.Runs;
            if (runs.Count == 0) {
                targetParagraph.Text = sourceParagraph.Text;
            } else {
                foreach (PowerPointTextRun run in runs) {
                    if (!options.IncludeBasicFormatting && HasBasicFormatting(run)) skippedBasicFormatting++;
                    Uri? hyperlink = run.Hyperlink;
                    bool clickActionRepresented = string.IsNullOrWhiteSpace(run.ClickAction)
                        || string.Equals(run.ClickAction, "ppaction://hlinksldjump", StringComparison.OrdinalIgnoreCase);
                    if (run.HasClickInteraction && (hyperlink == null || !clickActionRepresented
                        || run.ClickSoundName != null || run.ClickStopsSound)) {
                        unsupportedRunInteractions++;
                    }
                    if (run.HasMouseOverInteraction) unsupportedRunInteractions++;
                    if (hyperlink != null) {
                        ApplyPowerPointRun(run, targetParagraph.AddHyperlink(run.Text, hyperlink.ToString()), options);
                        if (!string.IsNullOrWhiteSpace(run.HyperlinkTooltip)) unsupportedHyperlinkTooltips++;
                    } else {
                        ApplyPowerPointRun(run, targetParagraph.AddRun(run.Text), options);
                    }
                    textRuns++;
                }
            }
            if (sourceParagraph.BulletCharacter != null || sourceParagraph.IsNumbered) listParagraphs++;
            paragraphs++;
        }
    }

    private static void CopyOdpParagraphsToPowerPoint(
        IReadOnlyList<OdpParagraph> sourceParagraphs,
        Func<IEnumerable<string>, IReadOnlyList<PowerPointParagraph>> setParagraphs,
        IReadOnlyList<OdpSlide> slides,
        ICollection<(PowerPointTextRun Run, int SlideIndex)> pendingInternalLinks,
        PowerPointOpenDocumentConversionOptions options,
        ref int paragraphs,
        ref int textRuns,
        ref int hyperlinks,
        ref int externalHyperlinks,
        ref int unsupportedHyperlinks,
        ref int unsupportedHyperlinkBehaviors,
        ref int approximatedRuns,
        ref int skippedBasicFormatting,
        ref int unsupportedWritingModes,
        ref int unsupportedMeasurements,
        ref int approximatedFontFamilyLists,
        ref int unsupportedFontFamilies) {
        IReadOnlyList<PowerPointParagraph> targetParagraphs =
            setParagraphs(sourceParagraphs.Select(paragraph => paragraph.Text));
        int paragraphCount = Math.Min(sourceParagraphs.Count, targetParagraphs.Count);
        for (int index = 0; index < paragraphCount; index++) {
            OdpParagraph sourceParagraph = sourceParagraphs[index];
            PowerPointParagraph targetParagraph = targetParagraphs[index];
            unsupportedMeasurements += ApplyOdpParagraphLayout(
                sourceParagraph,
                targetParagraph,
                ref unsupportedWritingModes);
            IReadOnlyList<OdpInlineNode> inlineNodes = sourceParagraph.InlineNodes;
            IReadOnlyList<PowerPointTextRun> existingRuns = targetParagraph.Runs;
            bool useExistingRun = existingRuns.Count > 0;

            PowerPointTextRun AddInlineRun(string text) {
                PowerPointTextRun result = useExistingRun
                    ? existingRuns[0]
                    : targetParagraph.AddRun(string.Empty);
                useExistingRun = false;
                result.Text = text;
                return result;
            }

            if (inlineNodes.Count == 0) {
                PowerPointTextRun run = useExistingRun
                    ? existingRuns[0]
                    : targetParagraph.AddRun(string.Empty);
                unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, run, options,
                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
            } else {
                foreach (OdpInlineNode node in inlineNodes) {
                    PowerPointTextRun targetRun = AddInlineRun(node.Text);
                    if (node.Kind == OdpInlineNodeKind.Run) {
                        unsupportedMeasurements += ApplyOdpRun(node.Run!, sourceParagraph, targetRun, options,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        if (!options.IncludeBasicFormatting && HasBasicFormatting(node.Run!)) skippedBasicFormatting++;
                    } else if (node.Kind == OdpInlineNodeKind.Hyperlink) {
                        OdpHyperlink hyperlink = node.Hyperlink!;
                        unsupportedMeasurements += ApplyOdpHyperlink(hyperlink, sourceParagraph, targetRun, options,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        if (!string.IsNullOrWhiteSpace(hyperlink.TargetFrameName)
                            || !string.IsNullOrWhiteSpace(hyperlink.ShowBehavior)) {
                            unsupportedHyperlinkBehaviors++;
                        }
                        if (TryResolveSlideFragment(hyperlink.Href, slides, out int targetSlideIndex)) {
                            pendingInternalLinks.Add((targetRun, targetSlideIndex));
                            hyperlinks++;
                        } else if (hyperlink.Href.StartsWith("#", StringComparison.Ordinal)) {
                            unsupportedHyperlinks++;
                        } else if (Uri.TryCreate(hyperlink.Href, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                            targetRun.SetHyperlink(uri);
                            hyperlinks++;
                            if (IsExternalOdfHref(hyperlink.Href)) externalHyperlinks++;
                        } else {
                            unsupportedHyperlinks++;
                        }
                        if (!options.IncludeBasicFormatting && HasBasicFormatting(hyperlink)) skippedBasicFormatting++;
                    } else {
                        unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, targetRun, options,
                            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        if (node.Kind == OdpInlineNodeKind.Other) approximatedRuns++;
                    }
                    textRuns++;
                }
            }
            if (!options.IncludeBasicFormatting && HasBasicFormatting(sourceParagraph)) skippedBasicFormatting++;
            paragraphs++;
        }
    }

    private static int ApplyOdpRun(OdpRun source, OdpParagraph paragraph, PowerPointTextRun target,
        PowerPointOpenDocumentConversionOptions options,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        target.Text = source.Text;
        if (!options.IncludeBasicFormatting) return 0;
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = source.Underline ?? paragraph.Underline ?? false;
        target.Strikethrough = source.StrikeThrough ?? paragraph.StrikeThrough ?? false;
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdpFontSize(fontSize, target);
        string? fontFamily = SelectOdfFontFamily(source.FontFamily ?? paragraph.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (fontFamily != null) target.FontName = fontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.Color = color.Value.ToString().TrimStart('#');
        OdfColor? background = source.BackgroundColor ?? paragraph.BackgroundColor;
        if (background.HasValue) target.HighlightColor = background.Value.ToString().TrimStart('#');
        return unsupported;
    }

    private static int ApplyOdpHyperlink(OdpHyperlink source, OdpParagraph paragraph,
        PowerPointTextRun target, PowerPointOpenDocumentConversionOptions options,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        if (!options.IncludeBasicFormatting) return 0;
        target.Bold = source.Bold ?? paragraph.Bold ?? false;
        target.Italic = source.Italic ?? paragraph.Italic ?? false;
        target.Underline = source.Underline ?? paragraph.Underline ?? false;
        target.Strikethrough = source.StrikeThrough ?? paragraph.StrikeThrough ?? false;
        OdfLength? fontSize = source.FontSize ?? paragraph.FontSize;
        int unsupported = ApplyOdpFontSize(fontSize, target);
        string? fontFamily = SelectOdfFontFamily(source.FontFamily ?? paragraph.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (fontFamily != null) target.FontName = fontFamily;
        OdfColor? color = source.Color ?? paragraph.Color;
        if (color.HasValue) target.Color = color.Value.ToString().TrimStart('#');
        OdfColor? background = source.BackgroundColor ?? paragraph.BackgroundColor;
        if (background.HasValue) target.HighlightColor = background.Value.ToString().TrimStart('#');
        return unsupported;
    }

    private static int ApplyOdpParagraphFormatting(OdpParagraph source, PowerPointTextRun target,
        PowerPointOpenDocumentConversionOptions options,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        if (!options.IncludeBasicFormatting) return 0;
        target.Bold = source.Bold == true;
        target.Italic = source.Italic == true;
        target.Underline = source.Underline == true;
        target.Strikethrough = source.StrikeThrough == true;
        int unsupported = ApplyOdpFontSize(source.FontSize, target);
        string? fontFamily = SelectOdfFontFamily(source.FontFamily,
            ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
        if (fontFamily != null) target.FontName = fontFamily;
        if (source.Color.HasValue) target.Color = source.Color.Value.ToString().TrimStart('#');
        if (source.BackgroundColor.HasValue) target.HighlightColor = source.BackgroundColor.Value.ToString().TrimStart('#');
        return unsupported;
    }

    private static int ApplyOdpFontSize(OdfLength? fontSize, PowerPointTextRun target) {
        if (!fontSize.HasValue) return 0;
        if (!fontSize.Value.TryToPoints(out double points)) return 1;
        target.FontSizePoints = points;
        return 0;
    }

    private static string? SelectOdfFontFamily(string? value,
        ref int approximatedFontFamilyLists, ref int unsupportedFontFamilies) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        if (!OdfFontFamilySyntax.TryParse(value, out OdfFontFamilySyntax? syntax)) {
            unsupportedFontFamilies++;
            return null;
        }
        if (syntax!.HasFallbacks) approximatedFontFamilyLists++;
        return syntax.PrimaryFamily;
    }
}
