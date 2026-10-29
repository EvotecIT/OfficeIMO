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
            int originalCount = document.Slides.Count;
            Action rollback = context.Deck.CaptureRollback();
            try {
                bool isFirst = content.IsFirst ?? originalCount == 0;
                PowerPointLayout layout = Resolve(context, content, isFirst, content.IsLast ?? false);
                var slides = new List<PowerPointSlide>();
                foreach (PowerPointLayoutSlide part in Expand(context, content, layout)) {
                    slides.Add(Render(context, part));
                }
                return slides;
            } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                while (document.Slides.Count > originalCount) document.RemoveSlide(document.Slides.Count - 1);
                rollback();
                throw;
            }
        }

        // ------------------------------------------------------------------ selection

        internal static PowerPointLayout Resolve(PowerPointLayoutContext context, PowerPointLayoutSlide slide, bool isFirst,
            bool isLast) {
            if (slide.Layout != PowerPointLayout.Auto) {
                if (Accepts(slide.Layout, slide)) return slide.Layout;
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.LayoutContentMismatch, OfficeDiagnosticSeverity.Warning,
                    "The " + slide.Layout + " layout does not accept this slide's content; the layout was chosen automatically.",
                    slide.Title);
            }

            return ResolveForPosition(slide, isFirst, isLast, slide.SelectionDeckSize ?? context.Document.Slides.Count + 1);
        }

        internal static PowerPointLayout ResolveForPosition(PowerPointLayoutSlide slide, bool isFirst, bool isLast, int deckSize) =>
            slide.Layout != PowerPointLayout.Auto && Accepts(slide.Layout, slide)
                ? slide.Layout : Select(slide, isFirst, isLast, deckSize);

        /// <summary>Automatic layout selection (powerpoint-layouts.json, selection.rules), first match wins.</summary>
        internal static PowerPointLayout Select(PowerPointLayoutSlide slide, bool isFirst, bool isLast, int deckSize) {
            List<PowerPointLayoutBlock> blocks = ContentBlocks(slide);
            if (isFirst && blocks.Count == 0 && !HasSemantic(slide)) return PowerPointLayout.Title;
            if (slide.Items<PowerPointLayoutLocation>().Any()) return PowerPointLayout.Coverage;
            bool kpis = slide.Items<PowerPointLayoutKpi>().Any();
            List<PowerPointLayoutCard> cards = slide.Items<PowerPointLayoutCard>().ToList();
            if (kpis && cards.Count > 0) return PowerPointLayout.CaseStudy;
            if (kpis) return PowerPointLayout.Kpi;
            if (cards.Count > 0) return cards.All(card => card.IsLogo) ? PowerPointLayout.Logos : PowerPointLayout.Cards;
            if (slide.Items<PowerPointLayoutStep>().Any()) return PowerPointLayout.Process;
            List<PowerPointLayoutColumn> columns = blocks.OfType<PowerPointLayoutColumn>().ToList();
            if (slide.Blocks.Count == 2 && blocks.Count == 2 && columns.Count == 2 && columns.All(c => !string.IsNullOrWhiteSpace(c.Title))) {
                return PowerPointLayout.Comparison;
            }

            if (slide.Blocks.Count == 1 && blocks.Count == 1 && blocks[0] is PowerPointLayoutImage image && IsFullBleedCandidate(image)) {
                return PowerPointLayout.ImageFull;
            }

            if (slide.Items<PowerPointLayoutQuote>().Any()) return PowerPointLayout.Quote;
            if (string.IsNullOrWhiteSpace(slide.Title) && blocks.Count == 1 && blocks[0] is PowerPointLayoutText text && text.Paragraphs.Count == 1 && text.Length <= 240) {
                return PowerPointLayout.Quote;
            }

            if (isLast && deckSize >= 5 && CanClose(slide) && (blocks.Count == 0 ||
                                            blocks.Count == 1 && blocks[0] is PowerPointLayoutList list && list.Items.Count <= 4 && list.Items.All(item => item.Level == 1))) {
                return PowerPointLayout.Closing;
            }

            if (blocks.Count == 0) return PowerPointLayout.Section;
            if (blocks.Count >= 2) return PowerPointLayout.TwoContent;
            return PowerPointLayout.TitleAndContent;
        }

        private static bool HasSemantic(PowerPointLayoutSlide slide) =>
            slide.Blocks.Any(b => b is PowerPointLayoutKpi or PowerPointLayoutCard or PowerPointLayoutStep or PowerPointLayoutLocation or PowerPointLayoutQuote);

        private static bool IsFullBleedCandidate(PowerPointLayoutImage image) {
            if (!PowerPointImageSize.TryRead(image.Path, out int width, out int height) || height == 0) return false;
            double aspect = width / (double)height;
            return width >= 1600 && Math.Abs(aspect / (16D / 9D) - 1D) <= 0.15D;
        }

        private static bool Accepts(PowerPointLayout layout, PowerPointLayoutSlide slide) {
            List<PowerPointLayoutBlock> blocks = ContentBlocks(slide);
            bool semantic = HasSemantic(slide);
            switch (layout) {
                case PowerPointLayout.Title:
                    return !semantic && blocks.Count == 0;
                case PowerPointLayout.Section:
                    return !semantic && blocks.Count == 0;
                case PowerPointLayout.TitleAndContent:
                    return !semantic && blocks.Count == 1;
                case PowerPointLayout.TwoContent:
                    return !semantic && blocks.Count >= 1;
                case PowerPointLayout.Comparison:
                    return slide.Blocks.Count == 2 && blocks.Count == 2 && blocks.All(b => b is PowerPointLayoutColumn);
                case PowerPointLayout.Kpi:
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutKpi);
                case PowerPointLayout.Cards:
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutCard);
                case PowerPointLayout.Logos:
                    // A logo may carry a short descriptor ("Authorized service") shown under its name.
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutCard card && card.Items.Count == 0 &&
                        (string.IsNullOrWhiteSpace(card.Text) || card.Text!.Length <= LogoCaptionLimit && card.Text.IndexOf('\n') < 0));
                case PowerPointLayout.Process:
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutStep);
                case PowerPointLayout.Quote:
                    return slide.Items<PowerPointLayoutQuote>().Any() || blocks.Count == 1 && blocks[0] is PowerPointLayoutText quote && quote.Paragraphs.Count == 1;
                case PowerPointLayout.ImageFull:
                    return slide.Blocks.Count == 1 && blocks.Count == 1 && blocks[0] is PowerPointLayoutImage;
                case PowerPointLayout.Closing:
                    return CanClose(slide);
                case PowerPointLayout.CaseStudy:
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutCard or PowerPointLayoutKpi) &&
                        slide.Items<PowerPointLayoutCard>().Count() <= PowerPointDeckPlanLimits.MaxCaseStudySections &&
                        slide.Items<PowerPointLayoutKpi>().Count() <= PowerPointDeckPlanLimits.MaxCaseStudyMetrics &&
                        slide.Items<PowerPointLayoutCard>().All(card => string.IsNullOrWhiteSpace(card.Image));
                case PowerPointLayout.Coverage:
                    return slide.Blocks.Count > 0 && slide.Blocks.All(block => block is PowerPointLayoutLocation);
                default:
                    return true;
            }
        }

        private const int LogoCaptionLimit = 40;

        private static bool CanClose(PowerPointLayoutSlide slide) {
            var (lede, blocks) = SplitLede(slide.Blocks);
            if (HasSemantic(slide) || lede != null && !string.IsNullOrWhiteSpace(slide.Subtitle)) return false;
            return blocks.Count == 0 || blocks.Count == 1 &&
                (blocks[0] is PowerPointLayoutList list && list.Items.Count <= 4 && list.Items.All(item => item.Level == 1) ||
                 blocks[0] is PowerPointLayoutText text && text.Paragraphs.Count == 1 && string.IsNullOrWhiteSpace(slide.Subtitle));
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
                case PowerPointLayout.Title: RenderTitle(context, writer, slide); break;
                case PowerPointLayout.Section: RenderSection(context, writer, slide); break;
                case PowerPointLayout.TwoContent: RenderTwoContent(context, writer, slide); break;
                case PowerPointLayout.Comparison: RenderComparison(context, writer, slide); break;
                case PowerPointLayout.Kpi: RenderKpi(context, writer, slide); break;
                case PowerPointLayout.Cards: RenderCards(context, writer, slide); break;
                case PowerPointLayout.Process: RenderProcess(context, writer, slide); break;
                case PowerPointLayout.Quote: RenderQuote(context, writer, slide); break;
                case PowerPointLayout.ImageFull: RenderImageFull(context, writer, slide); break;
                case PowerPointLayout.Closing: RenderClosing(context, writer, slide); break;
                case PowerPointLayout.CaseStudy: RenderCaseStudy(context, writer, slide); break;
                case PowerPointLayout.Logos: RenderLogos(context, writer, slide); break;
                case PowerPointLayout.Coverage: RenderCoverage(context, writer, slide); break;
                default: RenderTitleAndContent(context, writer, slide); break;
            }

            RenderFooter(context, writer, slide);
            if (!string.IsNullOrWhiteSpace(slide.Notes)) target.Notes.Text = slide.Notes!;
            return target;
        }

        private static PowerPointSlide AddNativeSlide(PowerPointLayoutContext context, PowerPointLayout layout) {
            PowerPointSlideLayoutInfo? selected = SelectNativeLayout(context, layout);
            return selected.HasValue ? context.Document.AddSlide(selected.Value.MasterIndex, selected.Value.LayoutIndex) : context.Document.AddSlide();
        }

        /// <summary>One layout selection owns both preflight geometry and physical slide creation.</summary>
        private static PowerPointSlideLayoutInfo? SelectNativeLayout(PowerPointLayoutContext context, PowerPointLayout layout, bool reportMissing = true) {
            string name;
            PowerPointSlideLayoutType[] types;
            switch (layout) {
                case PowerPointLayout.Title: name = "Title Slide"; types = new[] { PowerPointSlideLayoutType.Title }; break;
                case PowerPointLayout.Section: name = "Section Header"; types = new[] { PowerPointSlideLayoutType.SectionHeader }; break;
                case PowerPointLayout.TitleAndContent:
                    name = "Title and Content";
                    types = new[] { PowerPointSlideLayoutType.Object, PowerPointSlideLayoutType.Text };
                    break;
                case PowerPointLayout.TwoContent:
                    name = "Two Content";
                    types = new[] { PowerPointSlideLayoutType.TwoObjects, PowerPointSlideLayoutType.TwoColumnText };
                    break;
                case PowerPointLayout.Comparison:
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

            if (reportMissing && match < 0 && !context.Document.OwnsLayoutTheme) {
                context.Deck.Report(PowerPointLayoutDiagnosticCodes.TemplateLayoutMissing, OfficeDiagnosticSeverity.Information,
                    "The template has no layout for " + layout + "; its title-only layout is used instead.", oncePerDeck: true);
            }

            if (match < 0) match = layouts.FindIndex(l => l.Type == PowerPointSlideLayoutType.TitleOnly);
            if (match < 0) match = layouts.FindIndex(l => l.Type == PowerPointSlideLayoutType.Blank);
            if (match < 0) {
                if (reportMissing) context.Deck.Report(PowerPointLayoutDiagnosticCodes.TemplateLayoutMissing, OfficeDiagnosticSeverity.Information,
                    "The presentation has no layout for " + layout + "; the first layout was used.", oncePerDeck: true);
                return layouts.Count == 0 ? (PowerPointSlideLayoutInfo?)null : layouts[0];
            }

            return layouts[match];
        }
    }
}
