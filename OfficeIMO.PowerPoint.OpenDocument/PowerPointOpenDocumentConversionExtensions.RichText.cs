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
        ref int approximatedRuns,
        ref int skippedBasicFormatting,
        ref int unsupportedMeasurements) {
        IReadOnlyList<PowerPointParagraph> targetParagraphs =
            setParagraphs(sourceParagraphs.Select(paragraph => paragraph.Text));
        int paragraphCount = Math.Min(sourceParagraphs.Count, targetParagraphs.Count);
        for (int index = 0; index < paragraphCount; index++) {
            OdpParagraph sourceParagraph = sourceParagraphs[index];
            PowerPointParagraph targetParagraph = targetParagraphs[index];
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
                unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, run, options);
            } else {
                foreach (OdpInlineNode node in inlineNodes) {
                    PowerPointTextRun targetRun = AddInlineRun(node.Text);
                    if (node.Kind == OdpInlineNodeKind.Run) {
                        unsupportedMeasurements += ApplyOdpRun(node.Run!, sourceParagraph, targetRun, options);
                        if (!options.IncludeBasicFormatting && HasBasicFormatting(node.Run!)) skippedBasicFormatting++;
                    } else if (node.Kind == OdpInlineNodeKind.Hyperlink) {
                        OdpHyperlink hyperlink = node.Hyperlink!;
                        unsupportedMeasurements += ApplyOdpHyperlink(hyperlink, sourceParagraph, targetRun, options);
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
                        unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, targetRun, options);
                        if (node.Kind == OdpInlineNodeKind.Other) approximatedRuns++;
                    }
                    textRuns++;
                }
            }
            if (!options.IncludeBasicFormatting && HasBasicFormatting(sourceParagraph)) skippedBasicFormatting++;
            paragraphs++;
        }
    }
}
