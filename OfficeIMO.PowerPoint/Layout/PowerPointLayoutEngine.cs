using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     The slide layout system: selects a designed layout from the slide's content, resolves overflow into
    ///     continuation slides and renders each slide on the 12-column grid with theme tokens.
    /// </summary>
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Adds one logical slide; overflow can produce several physical slides.</summary>
        internal static IReadOnlyList<PowerPointSlide> AddSlides(PowerPointPresentation document, PowerPointLayoutSlide content,
            PowerPointLayoutTheme? theme = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (content == null) throw new ArgumentNullException(nameof(content));
            var context = new PowerPointLayoutContext(document, theme ?? document.LayoutTheme);
            bool isFirst = content.IsFirst ?? document.Slides.Count == 0;
            PowerPointLayoutKind layout = Resolve(context, content, isFirst, content.IsLast ?? false);
            var slides = new List<PowerPointSlide>();
            foreach (PowerPointLayoutSlide part in Expand(context, content, layout)) {
                slides.Add(Render(context, part));
            }

            return slides;
        }

        // ------------------------------------------------------------------ selection

        internal static PowerPointLayoutKind Resolve(PowerPointLayoutContext context, PowerPointLayoutSlide slide, bool isFirst,
            bool isLast) {
            if (slide.Layout != PowerPointLayoutKind.Auto) {
                if (Accepts(slide.Layout, slide)) return slide.Layout;
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.LayoutContentMismatch, OfficeDiagnosticSeverity.Warning,
                    "The " + slide.Layout + " layout does not accept this slide's content; the layout was chosen automatically.",
                    slide.Title);
            }

            return Select(slide, isFirst, isLast, context.Document.Slides.Count + 1);
        }

        /// <summary>Automatic layout selection (powerpoint-layouts.json, selection.rules), first match wins.</summary>
        internal static PowerPointLayoutKind Select(PowerPointLayoutSlide slide, bool isFirst, bool isLast, int deckSize) {
            List<PowerPointLayoutBlock> blocks = ContentBlocks(slide);
            if (isFirst && blocks.Count == 0 && !HasSemantic(slide)) return PowerPointLayoutKind.Title;
            if (slide.Items<PowerPointLayoutLocation>().Any()) return PowerPointLayoutKind.Coverage;
            bool kpis = slide.Items<PowerPointLayoutKpi>().Any();
            List<PowerPointLayoutCard> cards = slide.Items<PowerPointLayoutCard>().ToList();
            if (kpis && cards.Count > 0) return PowerPointLayoutKind.CaseStudy;
            if (kpis) return PowerPointLayoutKind.Kpi;
            if (cards.Count > 0) return cards.All(card => card.IsLogo) ? PowerPointLayoutKind.Logos : PowerPointLayoutKind.Cards;
            if (slide.Items<PowerPointLayoutStep>().Any()) return PowerPointLayoutKind.Process;
            List<PowerPointLayoutColumn> columns = blocks.OfType<PowerPointLayoutColumn>().ToList();
            if (blocks.Count == 2 && columns.Count == 2 && columns.All(c => !string.IsNullOrWhiteSpace(c.Title))) {
                return PowerPointLayoutKind.Comparison;
            }

            if (blocks.Count == 1 && blocks[0] is PowerPointLayoutImage image && IsFullBleedCandidate(image)) {
                return PowerPointLayoutKind.ImageFull;
            }

            if (slide.Items<PowerPointLayoutQuote>().Any()) return PowerPointLayoutKind.Quote;
            if (string.IsNullOrWhiteSpace(slide.Title) && blocks.Count == 1 && blocks[0] is PowerPointLayoutText text && text.Length <= 240) {
                return PowerPointLayoutKind.Quote;
            }

            if (isLast && deckSize >= 5 && (blocks.Count == 0 ||
                                            blocks.Count == 1 && blocks[0] is PowerPointLayoutList list && list.Items.Count <= 4)) {
                return PowerPointLayoutKind.Closing;
            }

            if (blocks.Count == 0) return PowerPointLayoutKind.Section;
            if (blocks.Count >= 2) return PowerPointLayoutKind.TwoContent;
            return PowerPointLayoutKind.TitleAndContent;
        }

        private static bool HasSemantic(PowerPointLayoutSlide slide) =>
            slide.Blocks.Any(b => b is PowerPointLayoutKpi or PowerPointLayoutCard or PowerPointLayoutStep or PowerPointLayoutLocation or PowerPointLayoutQuote);

        private static bool IsFullBleedCandidate(PowerPointLayoutImage image) {
            if (!PowerPointImageSize.TryRead(image.Path, out int width, out int height) || height == 0) return false;
            double aspect = width / (double)height;
            return width >= 1600 && Math.Abs(aspect / (16D / 9D) - 1D) <= 0.15D;
        }

        private static bool Accepts(PowerPointLayoutKind layout, PowerPointLayoutSlide slide) {
            List<PowerPointLayoutBlock> blocks = ContentBlocks(slide);
            bool semantic = HasSemantic(slide);
            switch (layout) {
                case PowerPointLayoutKind.Title:
                    return !semantic && blocks.Count <= 1;
                case PowerPointLayoutKind.Section:
                    return !semantic && blocks.Count == 0;
                case PowerPointLayoutKind.TitleAndContent:
                case PowerPointLayoutKind.TwoContent:
                    return !semantic && blocks.Count >= 1;
                case PowerPointLayoutKind.Comparison:
                    return blocks.Count == 2 && blocks.All(b => b is PowerPointLayoutColumn);
                case PowerPointLayoutKind.Kpi:
                    return slide.Items<PowerPointLayoutKpi>().Any();
                case PowerPointLayoutKind.Cards:
                    return slide.Items<PowerPointLayoutCard>().Any();
                case PowerPointLayoutKind.Logos:
                    return slide.Items<PowerPointLayoutCard>().Any();
                case PowerPointLayoutKind.Process:
                    return slide.Items<PowerPointLayoutStep>().Any();
                case PowerPointLayoutKind.Quote:
                    return slide.Items<PowerPointLayoutQuote>().Any() || blocks.Count == 1 && blocks[0] is PowerPointLayoutText;
                case PowerPointLayoutKind.ImageFull:
                    return blocks.Count == 1 && blocks[0] is PowerPointLayoutImage;
                case PowerPointLayoutKind.Closing:
                    return !semantic && (blocks.Count == 0 || blocks.Count == 1 && blocks[0] is PowerPointLayoutList or PowerPointLayoutText);
                case PowerPointLayoutKind.CaseStudy:
                    return slide.Items<PowerPointLayoutCard>().Any() || slide.Items<PowerPointLayoutKpi>().Any();
                case PowerPointLayoutKind.Coverage:
                    return slide.Items<PowerPointLayoutLocation>().Any();
                default:
                    return true;
            }
        }

        /// <summary>Content blocks counted by selection; drawings count as one visual block.</summary>
        internal static List<PowerPointLayoutBlock> ContentBlocks(PowerPointLayoutSlide slide) => SplitLede(slide.Blocks).Blocks;

        // ------------------------------------------------------------------ rendering dispatch

        private static PowerPointSlide Render(PowerPointLayoutContext context, PowerPointLayoutSlide slide) {
            PowerPointSlide target = slide.TemplateLayout != null
                ? context.Document.AddSlide(slide.TemplateLayout)
                : AddNativeSlide(context, slide.Layout);
            var writer = new PowerPointLayoutWriter(target, context.Theme);
            switch (slide.Layout) {
                case PowerPointLayoutKind.Title: RenderTitle(context, writer, slide); break;
                case PowerPointLayoutKind.Section: RenderSection(context, writer, slide); break;
                case PowerPointLayoutKind.TwoContent: RenderTwoContent(context, writer, slide); break;
                case PowerPointLayoutKind.Comparison: RenderComparison(context, writer, slide); break;
                case PowerPointLayoutKind.Kpi: RenderKpi(context, writer, slide); break;
                case PowerPointLayoutKind.Cards: RenderCards(context, writer, slide); break;
                case PowerPointLayoutKind.Process: RenderProcess(context, writer, slide); break;
                case PowerPointLayoutKind.Quote: RenderQuote(context, writer, slide); break;
                case PowerPointLayoutKind.ImageFull: RenderImageFull(context, writer, slide); break;
                case PowerPointLayoutKind.Closing: RenderClosing(context, writer, slide); break;
                case PowerPointLayoutKind.CaseStudy: RenderCaseStudy(context, writer, slide); break;
                case PowerPointLayoutKind.Logos: RenderLogos(context, writer, slide); break;
                case PowerPointLayoutKind.Coverage: RenderCoverage(context, writer, slide); break;
                default: RenderTitleAndContent(context, writer, slide); break;
            }

            RenderFooter(context, writer, slide);
            if (!string.IsNullOrWhiteSpace(slide.Notes)) target.Notes.Text = slide.Notes!;
            return target;
        }

        private static PowerPointSlide AddNativeSlide(PowerPointLayoutContext context, PowerPointLayoutKind layout) {
            string name;
            PowerPointSlideLayoutType[] types;
            switch (layout) {
                case PowerPointLayoutKind.Title: name = "Title Slide"; types = new[] { PowerPointSlideLayoutType.Title }; break;
                case PowerPointLayoutKind.Section: name = "Section Header"; types = new[] { PowerPointSlideLayoutType.SectionHeader }; break;
                case PowerPointLayoutKind.TitleAndContent:
                    name = "Title and Content";
                    types = new[] { PowerPointSlideLayoutType.Object, PowerPointSlideLayoutType.Text };
                    break;
                case PowerPointLayoutKind.TwoContent:
                    name = "Two Content";
                    types = new[] { PowerPointSlideLayoutType.TwoObjects, PowerPointSlideLayoutType.TwoColumnText };
                    break;
                case PowerPointLayoutKind.Comparison:
                    name = "Comparison";
                    types = new[] { PowerPointSlideLayoutType.TwoTextAndTwoObjects, PowerPointSlideLayoutType.TwoObjects };
                    break;
                default: name = "Title Only"; types = new[] { PowerPointSlideLayoutType.TitleOnly }; break;
            }

            PowerPointPresentation document = context.Document;
            List<PowerPointSlideLayoutInfo> layouts = document.GetSlideLayouts().ToList();
            int match = layouts.FindIndex(l => string.Equals(l.Name, "OfficeIMO " + layout, StringComparison.OrdinalIgnoreCase));
            if (match < 0) match = layouts.FindIndex(l => string.Equals(l.Name, layout.ToString(), StringComparison.OrdinalIgnoreCase));
            if (match < 0) match = layouts.FindIndex(l => string.Equals(l.Name, name, StringComparison.OrdinalIgnoreCase));
            foreach (PowerPointSlideLayoutType type in types) {
                if (match < 0) match = layouts.FindIndex(l => l.Type == type);
            }

            if (match < 0 && !context.Document.OwnsLayoutTheme) {
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.TemplateLayoutMissing, OfficeDiagnosticSeverity.Information,
                    "The template has no layout for " + layout + "; its title-only layout is used instead.", oncePerDeck: true);
            }

            if (match < 0) match = layouts.FindIndex(l => l.Type == PowerPointSlideLayoutType.TitleOnly);
            if (match < 0) match = layouts.FindIndex(l => l.Type == PowerPointSlideLayoutType.Blank);
            if (match < 0) {
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.TemplateLayoutMissing, OfficeDiagnosticSeverity.Information,
                    "The presentation has no layout for " + layout + "; the first layout was used.", oncePerDeck: true);
                return document.AddSlide();
            }

            return document.AddSlide(layouts[match].MasterIndex, layouts[match].LayoutIndex);
        }
    }
}
