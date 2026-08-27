using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    private sealed class PowerPointToOdpTextConversionState {
        internal int Paragraphs;
        internal int TextRuns;
        internal int LineBreaks;
        internal int Fields;
        internal int ListParagraphs;
        internal int ApproximatedAlignments;
        internal int SkippedBasicFormatting;
        internal int UnsupportedHyperlinkTooltips;
        internal int UnsupportedRunInteractions;
        internal int ApproximatedTextDecorations;
    }

    private static void CopyPowerPointTableCellToOdp(
        PowerPointTableCell source,
        OdpTableCell target,
        PowerPointOpenDocumentConversionOptions options,
        PowerPointToOdpTextConversionState state) {
        OdpParagraph? firstParagraph = target.Paragraphs.FirstOrDefault();
        bool useFirstParagraph = firstParagraph != null;

        OdpParagraph AddParagraph() {
            if (!useFirstParagraph) return target.AddParagraph();
            useFirstParagraph = false;
            firstParagraph!.Text = string.Empty;
            return firstParagraph;
        }

        CopyPowerPointParagraphsToOdp(source.Paragraphs, AddParagraph, options, state);
    }

    private static void CopyPowerPointParagraphsToOdp(
        IReadOnlyList<PowerPointParagraph> sourceParagraphs,
        Func<OdpParagraph> addParagraph,
        PowerPointOpenDocumentConversionOptions options,
        PowerPointToOdpTextConversionState state) {
        foreach (PowerPointParagraph sourceParagraph in sourceParagraphs) {
            OdpParagraph targetParagraph = addParagraph();
            if (ApplyPowerPointParagraphLayout(sourceParagraph, targetParagraph)) {
                state.ApproximatedAlignments++;
            }
            IReadOnlyList<PowerPointParagraphInline> inlineNodes = sourceParagraph.InlineNodes;
            if (inlineNodes.Count == 0) {
                targetParagraph.Text = sourceParagraph.Text;
            } else {
                foreach (PowerPointParagraphInline inlineNode in inlineNodes) {
                    if (inlineNode.Kind == PowerPointParagraphInlineKind.LineBreak) {
                        targetParagraph.AddText("\n");
                        state.LineBreaks++;
                        continue;
                    }
                    if (inlineNode.Kind == PowerPointParagraphInlineKind.Field) {
                        targetParagraph.AddText(inlineNode.Text);
                        state.Fields++;
                        continue;
                    }
                    PowerPointTextRun run = inlineNode.Run!;
                    if (HasApproximatedPowerPointUnderline(run.UnderlineStyle)) {
                        state.ApproximatedTextDecorations++;
                    }
                    if (!options.IncludeBasicFormatting && HasBasicFormatting(run)) state.SkippedBasicFormatting++;
                    Uri? hyperlink = run.Hyperlink;
                    bool clickActionRepresented = string.IsNullOrWhiteSpace(run.ClickAction)
                        || string.Equals(run.ClickAction, "ppaction://hlinksldjump", StringComparison.OrdinalIgnoreCase);
                    if (run.HasClickInteraction && (hyperlink == null || !clickActionRepresented
                        || run.ClickSoundName != null || run.ClickStopsSound)) {
                        state.UnsupportedRunInteractions++;
                    }
                    if (run.HasMouseOverInteraction) state.UnsupportedRunInteractions++;
                    if (hyperlink != null) {
                        ApplyPowerPointRun(run, targetParagraph.AddHyperlink(run.Text, hyperlink.ToString()), options);
                        if (!string.IsNullOrWhiteSpace(run.HyperlinkTooltip)) state.UnsupportedHyperlinkTooltips++;
                    } else {
                        ApplyPowerPointRun(run, targetParagraph.AddRun(run.Text), options);
                    }
                    state.TextRuns++;
                }
            }
            if (sourceParagraph.BulletCharacter != null || sourceParagraph.IsNumbered) state.ListParagraphs++;
            state.Paragraphs++;
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
        ref int approximatedParagraphAlignments,
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
                ref unsupportedWritingModes,
                ref approximatedParagraphAlignments);
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

            IReadOnlyList<PowerPointTextRun> AddInlineRuns(string text) {
                string[] segments = (text ?? string.Empty)
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n')
                    .Split('\n');
                var runs = new List<PowerPointTextRun>(segments.Length);
                for (int segmentIndex = 0; segmentIndex < segments.Length; segmentIndex++) {
                    if (segmentIndex > 0) targetParagraph.AddLineBreak();
                    if (segments[segmentIndex].Length > 0 || segments.Length == 1) {
                        runs.Add(AddInlineRun(segments[segmentIndex]));
                    }
                }
                return runs;
            }

            if (inlineNodes.Count == 0) {
                PowerPointTextRun run = useExistingRun
                    ? existingRuns[0]
                    : targetParagraph.AddRun(string.Empty);
                unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, run, options,
                    ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
            } else {
                foreach (OdpInlineNode node in inlineNodes) {
                    IReadOnlyList<PowerPointTextRun> targetRuns = AddInlineRuns(node.Text);
                    if (node.Kind == OdpInlineNodeKind.Run) {
                        foreach (PowerPointTextRun targetRun in targetRuns) {
                            unsupportedMeasurements += ApplyOdpRun(node.Run!, sourceParagraph, targetRun, options,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        }
                        if (!options.IncludeBasicFormatting && HasBasicFormatting(node.Run!)) skippedBasicFormatting++;
                    } else if (node.Kind == OdpInlineNodeKind.Hyperlink) {
                        OdpHyperlink hyperlink = node.Hyperlink!;
                        foreach (PowerPointTextRun targetRun in targetRuns) {
                            unsupportedMeasurements += ApplyOdpHyperlink(hyperlink, sourceParagraph, targetRun, options,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        }
                        if (!string.IsNullOrWhiteSpace(hyperlink.TargetFrameName)
                            || !string.IsNullOrWhiteSpace(hyperlink.ShowBehavior)) {
                            unsupportedHyperlinkBehaviors++;
                        }
                        if (TryResolveSlideFragment(hyperlink.Href, slides, out int targetSlideIndex)) {
                            foreach (PowerPointTextRun targetRun in targetRuns) {
                                pendingInternalLinks.Add((targetRun, targetSlideIndex));
                            }
                            hyperlinks++;
                        } else if (hyperlink.Href.StartsWith("#", StringComparison.Ordinal)) {
                            unsupportedHyperlinks++;
                        } else if (Uri.TryCreate(hyperlink.Href, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                            foreach (PowerPointTextRun targetRun in targetRuns) targetRun.SetHyperlink(uri);
                            hyperlinks++;
                            if (IsExternalOdfHref(hyperlink.Href)) externalHyperlinks++;
                        } else {
                            unsupportedHyperlinks++;
                        }
                        if (!options.IncludeBasicFormatting && HasBasicFormatting(hyperlink)) skippedBasicFormatting++;
                    } else {
                        foreach (PowerPointTextRun targetRun in targetRuns) {
                            unsupportedMeasurements += ApplyOdpParagraphFormatting(sourceParagraph, targetRun, options,
                                ref approximatedFontFamilyLists, ref unsupportedFontFamilies);
                        }
                        if (node.Kind == OdpInlineNodeKind.Other) approximatedRuns++;
                    }
                    textRuns += targetRuns.Count;
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
        ApplyOdpRunSemantics(
            source.Underline ?? paragraph.Underline,
            source.UnderlineStyle ?? paragraph.UnderlineStyle,
            source.UnderlineType ?? paragraph.UnderlineType,
            source.StrikeThrough ?? paragraph.StrikeThrough,
            source.LineThroughStyle ?? paragraph.LineThroughStyle,
            source.LineThroughType ?? paragraph.LineThroughType,
            source.TextPosition ?? paragraph.TextPosition,
            source.TextTransform ?? paragraph.TextTransform,
            source.SmallCaps ?? paragraph.SmallCaps,
            target);
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
        ApplyOdpRunSemantics(
            source.Underline ?? paragraph.Underline,
            source.UnderlineStyle ?? paragraph.UnderlineStyle,
            source.UnderlineType ?? paragraph.UnderlineType,
            source.StrikeThrough ?? paragraph.StrikeThrough,
            source.LineThroughStyle ?? paragraph.LineThroughStyle,
            source.LineThroughType ?? paragraph.LineThroughType,
            source.TextPosition ?? paragraph.TextPosition,
            source.TextTransform ?? paragraph.TextTransform,
            source.SmallCaps ?? paragraph.SmallCaps,
            target);
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
        ApplyOdpRunSemantics(
            source.Underline,
            source.UnderlineStyle,
            source.UnderlineType,
            source.StrikeThrough,
            source.LineThroughStyle,
            source.LineThroughType,
            source.TextPosition,
            source.TextTransform,
            source.SmallCaps,
            target);
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

    private static bool HasApproximatedPowerPointUnderline(PowerPointUnderlineStyle? style) => style is
        PowerPointUnderlineStyle.Words or
        PowerPointUnderlineStyle.Heavy or
        PowerPointUnderlineStyle.HeavyDotted or
        PowerPointUnderlineStyle.DashHeavy or
        PowerPointUnderlineStyle.DashLongHeavy or
        PowerPointUnderlineStyle.DotDashHeavy or
        PowerPointUnderlineStyle.DotDotDashHeavy or
        PowerPointUnderlineStyle.WavyHeavy;

    private static void ApplyOdpRunSemantics(
        bool? underline,
        OdfTextDecorationStyle? underlineStyle,
        OdfTextDecorationType? underlineType,
        bool? strike,
        OdfTextDecorationStyle? strikeStyle,
        OdfTextDecorationType? strikeType,
        OdfTextPosition? position,
        OdfTextTransform? transform,
        bool? smallCaps,
        PowerPointTextRun target) {
        target.UnderlineStyle = MapOdfUnderline(underline, underlineStyle, underlineType);
        bool hasStrike = strike == true && strikeStyle != OdfTextDecorationStyle.None && strikeType != OdfTextDecorationType.None;
        target.StrikeStyle = !hasStrike
            ? PowerPointStrikeStyle.None
            : strikeType == OdfTextDecorationType.Double
                ? PowerPointStrikeStyle.Double
                : PowerPointStrikeStyle.Single;
        target.BaselinePercent = position switch {
            OdfTextPosition.Superscript => 30D,
            OdfTextPosition.Subscript => -25D,
            OdfTextPosition.Normal => 0D,
            _ => null
        };
        target.Capitalization = transform == OdfTextTransform.Uppercase
            ? PowerPointCapitalization.AllCaps
            : smallCaps == true
                ? PowerPointCapitalization.SmallCaps
                : PowerPointCapitalization.None;
        if (transform == OdfTextTransform.Lowercase) target.TransformTextCase(OfficeTextCase.Lowercase);
        else if (transform == OdfTextTransform.Capitalize) target.TransformTextCase(OfficeTextCase.Capitalize);
    }

    private static PowerPointUnderlineStyle? MapOdfUnderline(
        bool? enabled,
        OdfTextDecorationStyle? style,
        OdfTextDecorationType? type) {
        if (enabled != true || style == OdfTextDecorationStyle.None || type == OdfTextDecorationType.None) {
            return PowerPointUnderlineStyle.None;
        }
        if (type == OdfTextDecorationType.Double) {
            return style == OdfTextDecorationStyle.Wave
                ? PowerPointUnderlineStyle.WavyDouble
                : PowerPointUnderlineStyle.Double;
        }
        return style switch {
            OdfTextDecorationStyle.Dotted => PowerPointUnderlineStyle.Dotted,
            OdfTextDecorationStyle.Dash => PowerPointUnderlineStyle.Dash,
            OdfTextDecorationStyle.LongDash => PowerPointUnderlineStyle.DashLong,
            OdfTextDecorationStyle.DotDash => PowerPointUnderlineStyle.DotDash,
            OdfTextDecorationStyle.DotDotDash => PowerPointUnderlineStyle.DotDotDash,
            OdfTextDecorationStyle.Wave => PowerPointUnderlineStyle.Wavy,
            _ => PowerPointUnderlineStyle.Single
        };
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
