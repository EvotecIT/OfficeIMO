using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     The designer slide kinds expressed as layout-system content. Every semantic slide is turned into a
    ///     <see cref="PowerPointLayoutSlide" /> and composed by <see cref="PowerPointLayoutEngine" />, so designer decks
    ///     and plain decks share one grid, one header, one type scale and the same overflow handling.
    /// </summary>
    internal static partial class PowerPointDesignExtensions {
        private const string LegacyFooterBrand = "OfficeIMO";
        private static readonly System.Text.RegularExpressions.Regex ContinuedTitle =
            new(@"^(.*?)( \(continued(?: \d+/\d+)?\))$", System.Text.RegularExpressions.RegexOptions.CultureInvariant);

        /// <remarks>
        ///     Layout runs reference the theme font slots (+mj-lt, +mn-lt), so text is measured with the fonts the
        ///     presentation's theme part names now. A template composed without applying the design keeps its own
        ///     fonts, and measuring with the design's fonts would wrap differently from what is drawn.
        /// </remarks>
        internal static PowerPointLayoutTheme LayoutThemeFor(PowerPointDocument presentation, PowerPointDesignTheme? theme) {
            if (theme == null) return presentation.LayoutTheme;
            // Loaded decks retain their native neutral palette, including readable muted text on dark slides.
            OfficeTheme basis = presentation.OwnsLayoutTheme ? presentation.BaseTheme : presentation.LayoutTheme.Theme;
            PowerPointLayoutTheme mapped = PowerPointLayoutTheme.From(ResolveTheme(theme), basis);
            if (presentation.OwnsLayoutTheme) return mapped;
            (string? heading, string? body) = presentation.ThemeFontFamilies();
            return mapped.WithFontFamilies(heading, body);
        }

        /// <summary>
        ///     Applies the designer theme: its brand accents and safe fonts are mapped onto the presentation's theme and
        ///     written to the theme part, and a presentation created by OfficeIMO also gets matching master text styles.
        /// </summary>
        public static PowerPointDocument ApplyDesignerTheme(this PowerPointDocument presentation,
            PowerPointDesignTheme? theme = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointLayoutTheme mapped = PowerPointLayoutTheme.From(resolvedTheme, presentation.BaseTheme);
            if (presentation.OwnsLayoutTheme) {
                presentation.ApplyLayoutTheme(mapped);
            } else {
                presentation.SetThemeColorsForAllMasters(new Dictionary<PowerPointThemeColor, string> {
                    [PowerPointThemeColor.Dark1] = mapped.Text,
                    [PowerPointThemeColor.Light1] = mapped.Background,
                    [PowerPointThemeColor.Dark2] = mapped.TextMuted,
                    [PowerPointThemeColor.Light2] = mapped.Surface,
                    [PowerPointThemeColor.Accent1] = mapped.Accent1,
                    [PowerPointThemeColor.Accent2] = mapped.Accent2,
                    [PowerPointThemeColor.Accent3] = mapped.Accent3,
                    [PowerPointThemeColor.Accent4] = mapped.Accent4,
                    [PowerPointThemeColor.Accent5] = mapped.Accent5,
                    [PowerPointThemeColor.Accent6] = mapped.Accent6
                });
                presentation.SetThemeFontsForAllMasters(new PowerPointThemeFontSet(mapped.HeadingFont, mapped.BodyFont, null, null, null, null));
            }

            presentation.ThemeName = resolvedTheme.Name;
            return presentation;
        }

        private static PowerPointLayoutSlide Content(PowerPointDocument presentation, string? title, string? subtitle,
            PowerPointDesignerSlideOptions options, PowerPointLayout layout) {
            PowerPointLayoutDeck deck = presentation.LayoutDeck;
            if (!string.IsNullOrWhiteSpace(options.FooterLeft) && options.FooterLeft != LegacyFooterBrand) {
                deck.FooterText ??= options.FooterLeft;
            }

            string? suffix = null;
            if (title != null) {
                System.Text.RegularExpressions.Match match = ContinuedTitle.Match(title);
                if (match.Success) {
                    title = match.Groups[1].Value;
                    suffix = match.Groups[2].Value;
                }
            }

            var slide = new PowerPointLayoutSlide {
                ContinuedSuffix = suffix,
                Title = title,
                Subtitle = string.IsNullOrWhiteSpace(subtitle) || suffix != null ? null : subtitle,
                Layout = layout,
                FooterRightText = options.FooterRight,
                TemplateLayout = options.TemplateLayout
            };
            // A section title becomes the eyebrow of the slides that follow it; the deck eyebrow fills in before the first section.
            if (string.IsNullOrWhiteSpace(deck.CurrentSection) && !string.IsNullOrWhiteSpace(options.Eyebrow)) {
                slide.Eyebrow = options.Eyebrow;
            }

            return slide;
        }

        private static PowerPointSlide Compose(PowerPointDocument presentation, PowerPointLayoutSlide slide, PowerPointDesignTheme? theme,
            string? title) {
            IReadOnlyList<PowerPointSlide> slides = PowerPointLayoutEngine.AddSlides(presentation, slide, LayoutThemeFor(presentation, theme));
            foreach (PowerPointSlide added in slides) FinalizeDesignerAccessibility(added, title);
            return slides[0];
        }

        private static void ReportIgnored(PowerPointDocument presentation, string option, string? title) {
            presentation.LayoutDeck.Report("designer-option-ignored", OfficeDiagnosticSeverity.Information,
                "The " + option + " designer option has no effect in the theme layouts; the slide uses its designed layout.", title,
                oncePerDeck: false);
        }

        // ------------------------------------------------------------------ section, case study, cards, process

        public static PowerPointSlide AddDesignerSectionSlide(this PowerPointDocument presentation, string title,
            string? subtitle = null, PowerPointDesignTheme? theme = null, PowerPointDesignerSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointDesignerSlideOptions resolved = options ?? new PowerPointDesignerSlideOptions();
            bool cover = presentation.Slides.Count == 0;
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle, resolved,
                cover ? PowerPointLayout.Title : PowerPointLayout.Section);
            if (cover) slide.Eyebrow = resolved.Eyebrow;
            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerCaseStudySlide(this PowerPointDocument presentation,
            string clientTitle, IEnumerable<PowerPointCaseStudySection> sections, IEnumerable<PowerPointMetric>? metrics = null,
            PowerPointDesignTheme? theme = null, PowerPointCaseStudySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (clientTitle == null) throw new ArgumentNullException(nameof(clientTitle));
            PowerPointCaseStudySlideOptions resolved = options ?? new PowerPointCaseStudySlideOptions();
            List<PowerPointCaseStudySection> sectionList = NormalizeSections(sections, 4, nameof(sections));
            List<PowerPointMetric> metricList = (metrics ?? Enumerable.Empty<PowerPointMetric>()).Where(m => m != null).ToList();
            if (metricList.Count > 3) {
                throw new ArgumentOutOfRangeException(nameof(metrics), "A case study shows up to three result metrics.");
            }
            PowerPointLayoutSlide slide = Content(presentation, clientTitle, null, resolved, PowerPointLayout.CaseStudy);
            foreach (PowerPointCaseStudySection section in sectionList) slide.Blocks.Add(new PowerPointLayoutCard(section.Heading, section.Body));
            foreach (PowerPointMetric metric in metricList.Take(3)) slide.Blocks.Add(Kpi(metric));
            if (!string.IsNullOrWhiteSpace(resolved.VisualImagePath) || !string.IsNullOrWhiteSpace(resolved.PersonImagePath)) {
                ReportIgnored(presentation, "case-study visual image", clientTitle);
            }

            return Compose(presentation, slide, theme, clientTitle);
        }

        public static PowerPointSlide AddDesignerCardGridSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointCardContent> cards,
            PowerPointDesignTheme? theme = null, PowerPointCardGridSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointCardGridSlideOptions resolved = options ?? new PowerPointCardGridSlideOptions();
            List<PowerPointCardContent> cardList = NormalizeCards(cards);
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Cards);
            foreach (PowerPointCardContent card in cardList) slide.Blocks.Add(Card(card));
            slide.Takeaway = resolved.SupportingText;
            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerProcessSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointProcessStep> steps,
            PowerPointDesignTheme? theme = null, PowerPointProcessSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointProcessSlideOptions resolved = options ?? new PowerPointProcessSlideOptions();
            List<PowerPointProcessStep> stepList = NormalizeSteps(steps);
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Process);
            foreach (PowerPointProcessStep step in stepList) slide.Blocks.Add(new PowerPointLayoutStep(step.Title, step.Body, step.Number));
            return Compose(presentation, slide, theme, title);
        }

        // ------------------------------------------------------------------ logos, coverage, capability

        public static PowerPointSlide AddDesignerLogoWallSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointLogoItem> logos,
            PowerPointDesignTheme? theme = null, PowerPointLogoWallSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointLogoWallSlideOptions resolved = options ?? new PowerPointLogoWallSlideOptions();
            List<PowerPointLogoItem> logoList = NormalizeLogoItems(logos);
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Logos);
            foreach (PowerPointLogoItem logo in logoList) {
                slide.Blocks.Add(new PowerPointLayoutCard(logo.Name, logo.Subtitle, image: logo.ImagePath, color: PowerPointLayoutBlock.ParseColor(logo.AccentColor)));
            }

            slide.Takeaway = resolved.SupportingText;
            if (!string.IsNullOrWhiteSpace(resolved.FeaturedImagePath)) ReportIgnored(presentation, "logo-wall featured image", title);
            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerCoverageSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointCoverageLocation> locations,
            PowerPointDesignTheme? theme = null, PowerPointCoverageSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointCoverageSlideOptions resolved = options ?? new PowerPointCoverageSlideOptions();
            List<PowerPointCoverageLocation> locationList = NormalizeLocations(locations);
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle ?? resolved.SupportingText, resolved, PowerPointLayout.Coverage);
            slide.MapLabel = resolved.MapLabel;
            foreach (PowerPointCoverageLocation location in locationList) {
                slide.Blocks.Add(new PowerPointLayoutLocation(location.Name, location.X, location.Y, location.Detail));
            }

            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerCapabilitySlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointCapabilitySection> sections,
            PowerPointDesignTheme? theme = null, PowerPointCapabilitySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (title == null) throw new ArgumentNullException(nameof(title));
            PowerPointCapabilitySlideOptions resolved = options ?? new PowerPointCapabilitySlideOptions();
            List<PowerPointCapabilitySection> sectionList = NormalizeCapabilitySections(sections);
            PowerPointLayoutSlide slide;
            if (resolved.Metrics.Count > 3) {
                throw new ArgumentOutOfRangeException(nameof(options), "A capability slide shows up to three result metrics.");
            }

            if (resolved.Metrics.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.CaseStudy);
                foreach (PowerPointCapabilitySection section in sectionList) slide.Blocks.Add(new PowerPointLayoutCard(section.Heading, section.Body, section.Items));
                foreach (PowerPointMetric metric in resolved.Metrics.Take(3)) slide.Blocks.Add(Kpi(metric));
            } else if (resolved.Locations.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Coverage);
                slide.MapLabel = resolved.VisualLabel;
                foreach (PowerPointCoverageLocation location in resolved.Locations) {
                    slide.Blocks.Add(new PowerPointLayoutLocation(location.Name, location.X, location.Y, location.Detail));
                }

                ReportIgnored(presentation, "capability sections beside a coverage visual", title);
            } else if (!string.IsNullOrWhiteSpace(resolved.VisualImagePath) && System.IO.File.Exists(resolved.VisualImagePath)) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TwoContent);
                slide.Blocks.Add(new PowerPointLayoutList(sectionList.SelectMany(SectionItems)));
                slide.Blocks.Add(new PowerPointLayoutImage(resolved.VisualImagePath!, resolved.VisualLabel ?? title));
            } else if (resolved.Logos.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Logos);
                foreach (PowerPointLogoItem logo in resolved.Logos) slide.Blocks.Add(new PowerPointLayoutCard(logo.Name, logo.Subtitle, image: logo.ImagePath));
                ReportIgnored(presentation, "capability sections beside a logo visual", title);
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Cards);
                foreach (PowerPointCapabilitySection section in sectionList) {
                    slide.Blocks.Add(new PowerPointLayoutCard(section.Heading, section.Body, section.Items, color: PowerPointLayoutBlock.ParseColor(section.AccentColor)));
                }
            }

            return Compose(presentation, slide, theme, title);
        }

        private static IEnumerable<PowerPointLayoutListItem> SectionItems(PowerPointCapabilitySection section) {
            yield return new PowerPointLayoutListItem(string.IsNullOrWhiteSpace(section.Body) ? section.Heading : section.Heading + ": " + section.Body);
            foreach (string item in section.Items) yield return new PowerPointLayoutListItem(item, 2);
        }

        // ------------------------------------------------------------------ executive, chart, comparison, table

        public static PowerPointSlide AddDesignerExecutiveSummarySlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointExecutiveSummaryContent content,
            PowerPointDesignTheme? theme = null, PowerPointExecutiveSummarySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointExecutiveSummarySlideOptions resolved = options ?? new PowerPointExecutiveSummarySlideOptions();
            PowerPointLayoutSlide slide;
            if (content.Metrics.Count > 0 && content.Points.Count > 0) {
                if (content.Metrics.Count > 3) {
                    throw new ArgumentException("An executive summary with decision points shows up to three result metrics; split the remaining metrics onto a Kpi slide.",
                        nameof(content));
                }

                slide = Content(presentation, title, subtitle ?? content.Lead, resolved, PowerPointLayout.CaseStudy);
                foreach (PowerPointCardContent point in content.Points) slide.Blocks.Add(new PowerPointLayoutCard(point.Title, null, point.Items));
                foreach (PowerPointMetric metric in content.Metrics.Take(3)) slide.Blocks.Add(Kpi(metric));
            } else if (content.Metrics.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Kpi);
                foreach (PowerPointMetric metric in content.Metrics) slide.Blocks.Add(Kpi(metric));
                slide.Takeaway = content.Lead;
            } else if (content.Points.Count > 0) {
                slide = Content(presentation, title, subtitle ?? content.Lead, resolved, PowerPointLayout.Cards);
                foreach (PowerPointCardContent point in content.Points) slide.Blocks.Add(Card(point));
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TitleAndContent);
                slide.Blocks.Add(new PowerPointLayoutText(content.Lead ?? string.Empty));
            }

            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerChartStorySlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointChartStoryContent content,
            PowerPointDesignTheme? theme = null, PowerPointChartStorySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointChartStorySlideOptions resolved = options ?? new PowerPointChartStorySlideOptions();
            var chart = new PowerPointLayoutChart(content.ChartKind, content.SharedData) {
                AlternativeText = content.AlternativeText ?? content.DataSummary
            };
            PowerPointLayoutSlide slide;
            if (content.Insights.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TwoContent);
                slide.Blocks.Add(new PowerPointLayoutList(content.Insights));
                slide.Blocks.Add(chart);
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TitleAndContent);
                if (!string.IsNullOrWhiteSpace(content.Caption)) slide.Blocks.Add(new PowerPointLayoutText(content.Caption!));
                slide.Blocks.Add(chart);
            }

            slide.SourceNote = Join(content.Insights.Count > 0 ? content.Caption : null, Source(content.Provenance));
            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerComparisonSlide(this PowerPointDocument presentation,
            string title, string? subtitle, IEnumerable<PowerPointComparisonItem> items,
            PowerPointDesignTheme? theme = null, PowerPointComparisonSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            List<PowerPointComparisonItem> values = (items ?? throw new ArgumentNullException(nameof(items))).ToList();
            if (values.Count < 2 || values.Count > 4) {
                throw new ArgumentOutOfRangeException(nameof(items), "Comparison slides support two to four options.");
            }

            PowerPointComparisonSlideOptions resolved = options ?? new PowerPointComparisonSlideOptions();
            PowerPointLayoutSlide slide;
            if (values.Count == 2) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Comparison);
                foreach (PowerPointComparisonItem item in values) {
                    var blocks = new List<PowerPointLayoutBlock>();
                    if (!string.IsNullOrWhiteSpace(item.Summary)) blocks.Add(new PowerPointLayoutText(item.Summary!));
                    blocks.Add(new PowerPointLayoutList(ComparisonItems(item)));
                    slide.Blocks.Add(new PowerPointLayoutColumn(item.Title, blocks));
                }
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.Cards);
                foreach (PowerPointComparisonItem item in values) {
                    slide.Blocks.Add(new PowerPointLayoutCard(item.Title, item.Summary, ComparisonItems(item).Select(i => i.Text)));
                }
            }

            return Compose(presentation, slide, theme, title);
        }

        private static IEnumerable<PowerPointLayoutListItem> ComparisonItems(PowerPointComparisonItem item) {
            foreach (string strength in item.Strengths) yield return new PowerPointLayoutListItem(strength);
            foreach (string tradeoff in item.Tradeoffs) yield return new PowerPointLayoutListItem("Trade-off: " + tradeoff);
        }

        public static PowerPointSlide AddDesignerAppendixTableSlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointTableData data, PowerPointDesignTheme? theme = null,
            PowerPointAppendixTableSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (data == null) throw new ArgumentNullException(nameof(data));
            PowerPointAppendixTableSlideOptions resolved = options ?? new PowerPointAppendixTableSlideOptions();
            var table = new PowerPointLayoutTable(data.Headers, data.Rows) { Caption = Join(data.Caption, Source(data.Provenance)) };
            PowerPointLayoutSlide slide;
            if (data.Notes.Count > 0) {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TwoContent);
                slide.Blocks.Add(table);
                slide.Blocks.Add(new PowerPointLayoutList(data.Notes));
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TitleAndContent);
                slide.Blocks.Add(table);
            }

            return Compose(presentation, slide, theme, title);
        }

        // ------------------------------------------------------------------ screenshot, architecture, closing

        public static PowerPointSlide AddDesignerScreenshotStorySlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointImageAsset image, IEnumerable<string>? narrative = null,
            PowerPointDesignTheme? theme = null, PowerPointScreenshotStorySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (image == null) throw new ArgumentNullException(nameof(image));
            image.Validate();
            PowerPointScreenshotStorySlideOptions resolved = options ?? new PowerPointScreenshotStorySlideOptions();
            List<string> points = (narrative ?? Array.Empty<string>()).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
            var picture = new PowerPointLayoutImage(image.Path, image.AlternativeText) { FocalX = image.FocalX, FocalY = image.FocalY };
            foreach (PowerPointImageAnnotation annotation in image.Annotations) {
                picture.Annotations.Add((annotation.X, annotation.Y, string.IsNullOrWhiteSpace(annotation.Detail) ? annotation.Label : annotation.Label + " \u2013 " + annotation.Detail));
            }
            PowerPointLayoutSlide slide;
            if (points.Count == 0) {
                slide = Content(presentation, title, subtitle, resolved,
                    image.Placement == PowerPointImagePlacement.Fill ? PowerPointLayout.ImageFull : PowerPointLayout.TitleAndContent);
                slide.Blocks.Add(picture);
            } else {
                slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TwoContent);
                slide.Blocks.Add(new PowerPointLayoutList(points));
                slide.Blocks.Add(picture);
            }

            slide.SourceNote = Join(image.Caption, Source(image.Provenance));

            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerArchitectureSlide(this PowerPointDocument presentation,
            string title, string? subtitle, PowerPointArchitectureContent content,
            PowerPointDesignTheme? theme = null, PowerPointArchitectureSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointArchitectureSlideOptions resolved = options ?? new PowerPointArchitectureSlideOptions();
            PowerPointLayoutSlide slide = Content(presentation, title, subtitle, resolved, PowerPointLayout.TitleAndContent);
            slide.Blocks.Add(new PowerPointLayoutDrawing((target, context, box) => PowerPointLayoutDiagram.Draw(target, context, box, content)));
            return Compose(presentation, slide, theme, title);
        }

        public static PowerPointSlide AddDesignerClosingSlide(this PowerPointDocument presentation,
            string title, PowerPointClosingContent content, PowerPointDesignTheme? theme = null,
            PowerPointClosingSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            PowerPointClosingSlideOptions resolved = options ?? new PowerPointClosingSlideOptions();
            PowerPointLayoutSlide slide = Content(presentation, title, Join(content.Statement, content.Contact, "\n"), resolved, PowerPointLayout.Closing);
            slide.Eyebrow = resolved.Eyebrow;
            if (!string.IsNullOrWhiteSpace(content.CallToAction)) slide.Blocks.Add(new PowerPointLayoutList(new[] { content.CallToAction! }));
            return Compose(presentation, slide, theme, title);
        }

        // ------------------------------------------------------------------ mapping helpers

        private static PowerPointLayoutKpi Kpi(PowerPointMetric metric) =>
            new(metric.Value, metric.Label, metric.Change, PowerPointLayoutBlock.ParseColor(metric.Color));

        private static PowerPointLayoutCard Card(PowerPointCardContent card) =>
            new(card.Title, card.Text, card.Items, card.ImagePath, PowerPointLayoutBlock.ParseColor(card.AccentColor));

        private static string? Source(string? provenance) =>
            string.IsNullOrWhiteSpace(provenance) ? null : "Source: " + provenance!.Trim();

        private static string? Join(string? first, string? second, string separator = " · ") {
            if (string.IsNullOrWhiteSpace(first)) return string.IsNullOrWhiteSpace(second) ? null : second;
            if (string.IsNullOrWhiteSpace(second)) return first;
            return first + separator + second;
        }
    }
}
