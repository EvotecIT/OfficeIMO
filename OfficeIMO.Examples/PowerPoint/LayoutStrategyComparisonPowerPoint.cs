using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates using the same semantic plan with different design brief variation controls.
    /// </summary>
    public static class LayoutStrategyComparisonPowerPoint {
        public static void Example_LayoutStrategyComparisonPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Layout strategy comparison");
            string filePath = Path.Join(folderPath, "PowerPoint Layout Strategy Comparison.pptx");

            PowerPointDeckPlan plan = CreatePlan();
            LayoutStrategyExample[] examples = {
                new(PowerPointAutoLayoutStrategy.ContentFirst, PowerPointPaletteStyle.SplitComplementary,
                    "Content fit first", "Keep dense or structured content readable before adding visual variety."),
                new(PowerPointAutoLayoutStrategy.Compact, PowerPointPaletteStyle.CoolNeutral,
                    "Compact business rhythm", "Prefer tighter layouts when the deck needs more information per slide."),
                new(PowerPointAutoLayoutStrategy.VisualFirst, PowerPointPaletteStyle.Complementary,
                    "Visual proof first", "Favor stronger hero, proof, and visual framing when the content allows it.")
            };

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

            foreach (LayoutStrategyExample example in examples) {
                PowerPointDesignBrief brief = CreateBrief(example);
                IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> preview = brief.DescribeDeckPlan(plan);
                PowerPointDeckComposer deck = presentation.UseDesigner(brief);

                deck.AddSectionSlide(example.Title,
                    example.Description,
                    "layout-strategy-" + example.Strategy,
                    options => options.SectionVariant = example.Strategy == PowerPointAutoLayoutStrategy.VisualFirst
                        ? PowerPointSectionLayoutVariant.Poster
                        : PowerPointSectionLayoutVariant.EditorialRail);

                deck.ComposeSlide(composer => {
                    composer.AddTitle("Resolved variants",
                        "The plan is unchanged; the brief steers palette, density, and Auto layout behavior.");
                    PowerPointLayoutBox[] columns = composer.ContentColumns(2, 0.8, topCm: 3.7, bottomMarginCm: 2.35);

                    composer.AddCardGrid(preview.Select((summary, index) => new PowerPointCardContent(
                            (index + 1) + ". " + ToDisplayName(summary.Kind.ToString()),
                            new[] {
                                summary.Title,
                                "Variant: " + ToDisplayName(summary.LayoutVariant ?? "custom"),
                                "Items: " + summary.ContentItemCount,
                            })),
                        columns[0],
                        new PowerPointCardGridSlideOptions {
                            MaxColumns = 2,
                            Variant = PowerPointCardGridLayoutVariant.SoftTiles
                        });

                    int distinctVariants = preview
                        .Select(summary => summary.LayoutVariant ?? "custom")
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Count();
                    composer.AddMetricStrip(new[] {
                        new PowerPointMetric(preview.Count.ToString(), "slides"),
                        new PowerPointMetric(distinctVariants.ToString(), "variants"),
                        new PowerPointMetric(((int)example.Strategy + 1).ToString(), "strategy")
                    }, columns[1].TakeTopCm(2.2));

                    string reasons = "Palette " + ToDisplayName(example.PaletteStyle.ToString()) + ". Mode " +
                        ToDisplayName(example.Strategy.ToString()) + ". Preview variants before rendering.";
                    composer.AddCalloutBand(reasons,
                        columns[1].InsetCm(0, 2.75, 0, 0).TakeTopCm(2.25));
                }, "layout-preview-" + example.Strategy,
                    options => options.FooterRight = ToDisplayName(example.Strategy.ToString()));

                deck.AddSlides(plan);
            }

            presentation.Save();
            Validate(filePath, presentation);
            Helpers.Open(filePath, openPowerPoint);
        }

        private static PowerPointDesignBrief CreateBrief(LayoutStrategyExample example) {
            return PowerPointDesignBrief
                .FromBrand("#008C95", "layout-strategy-comparison", "service portfolio and delivery proposal")
                .WithIdentity("Variation Controls", eyebrow: "OfficeIMO.PowerPoint", footerLeft: "OFFICEIMO",
                    footerRight: ToDisplayName(example.Strategy.ToString()))
                .WithPaletteStyle(example.PaletteStyle)
                .WithPalette(secondaryAccentColor: "#6D5BD0", tertiaryAccentColor: "#0E7490",
                    warmAccentColor: "#FFB000", surfaceColor: "#F6FAFC", panelBorderColor: "#D5E3EA")
                .WithLayoutStrategy(example.Strategy)
                .WithVariety(PowerPointDesignVariety.Exploratory)
                .WithPreferredMoods(PowerPointDesignMood.Energetic, PowerPointDesignMood.Editorial)
                .WithPreferredVisualStyles(PowerPointVisualStyle.Geometric, PowerPointVisualStyle.Soft);
        }

        private static PowerPointDeckPlan CreatePlan() {
            return new PowerPointDeckPlan()
                .AddSection("Service portfolio",
                    "One semantic plan can produce different slide rhythm without changing coordinates.",
                    "layout-cover")
                .AddCaseStudy("Managed workplace rollout",
                    new[] {
                        new PowerPointCaseStudySection("Client",
                            "A distributed organization needed a clearer service story."),
                        new PowerPointCaseStudySection("Challenge",
                            "Many locations, mixed hardware, and operational details had to stay readable."),
                        new PowerPointCaseStudySection("Solution",
                            "The rollout model separated discovery, delivery, reporting, and support ownership."),
                        new PowerPointCaseStudySection("Result",
                            "The same content can become a structured case study or a more visual proof slide.")
                    },
                    new[] {
                        new PowerPointMetric("18", "sites"),
                        new PowerPointMetric("420", "devices")
                    },
                    "layout-case")
                .AddProcess("Delivery path",
                    "Auto variants react to the strategy and the number of steps.",
                    new[] {
                        new PowerPointProcessStep("Discover", "Collect constraints and service expectations."),
                        new PowerPointProcessStep("Design", "Choose target architecture and rollout rules."),
                        new PowerPointProcessStep("Pilot", "Validate with a controlled group."),
                        new PowerPointProcessStep("Roll out", "Deliver in waves with reporting."),
                        new PowerPointProcessStep("Operate", "Move into repeatable support.")
                    },
                    "layout-process")
                .AddCardGrid("Scope of services",
                    "Cards remain semantic while the brief chooses their visual treatment.",
                    new[] {
                        new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot", "Policy baseline" }),
                        new PowerPointCardContent("Operations", new[] { "Incident handling", "Monitoring", "Optimization" }),
                        new PowerPointCardContent("Consulting", new[] { "Roadmap", "Architecture", "Discovery" }),
                        new PowerPointCardContent("Audits", new[] { "Configuration", "Security review", "Modernization plan" })
                    },
                    "layout-cards")
                .AddCoverage("Delivery coverage",
                    null,
                    new[] {
                        new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
                        new PowerPointCoverageLocation("Warszawa", 0.60, 0.48),
                        new PowerPointCoverageLocation("Poznan", 0.34, 0.46),
                        new PowerPointCoverageLocation("Wroclaw", 0.36, 0.70),
                        new PowerPointCoverageLocation("Krakow", 0.58, 0.78)
                    },
                    "layout-coverage");
        }

        private static void Validate(string filePath, PowerPointPresentation presentation) {
            List<ValidationErrorInfo> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string details = string.Join(Environment.NewLine, errors.Take(5).Select(error => error.Description));
                throw new InvalidOperationException($"PowerPoint validation failed with {errors.Count} error(s).{Environment.NewLine}{details}");
            }

            Console.WriteLine($"    Saved: {filePath}");
            Console.WriteLine("    Validation: no Open XML errors found.");
        }

        private static string ToDisplayName(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return value;
            }

            List<char> chars = new();
            for (int i = 0; i < value.Length; i++) {
                char current = value[i];
                if (i > 0 && char.IsUpper(current) &&
                    (char.IsLower(value[i - 1]) || (i + 1 < value.Length && char.IsLower(value[i + 1])))) {
                    chars.Add(' ');
                }

                chars.Add(current);
            }

            return new string(chars.ToArray());
        }

        private sealed record LayoutStrategyExample(PowerPointAutoLayoutStrategy Strategy,
            PowerPointPaletteStyle PaletteStyle, string Title, string Description);
    }
}
