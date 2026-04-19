using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates high-level designer compositions for visually stronger decks.
    /// </summary>
    public static class DesignerPowerPointDeck {
        public static void Example_DesignerPowerPointDeck(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Designer composition deck");
            string filePath = Path.Join(folderPath, "Designer PowerPoint Deck.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDesignRecipe recipe = PowerPointDesignRecipe.FindBuiltIn("consulting portfolio")
                ?? PowerPointDesignRecipe.ConsultingPortfolio;
            IReadOnlyList<PowerPointDeckDesign> alternatives = recipe.CreateAlternativesFromBrand(
                "#008C95", "designer-example", name: "OfficeIMO Teal", eyebrow: "OfficeIMO.PowerPoint",
                footerLeft: "OFFICEIMO", footerRight: "The Good Slides");
            PowerPointDeckDesign design = alternatives[0];
            PowerPointDeckComposer deck = presentation.UseDesigner(design);

            deck.AddSectionSlide("Case Study", "Project portfolio", "section",
                options => options.SectionVariant = PowerPointSectionLayoutVariant.Poster);

            deck.AddCaseStudySlide("Randstad - print environment rollout",
                new[] {
                    new PowerPointCaseStudySection("Client",
                        "A distributed organization needed one clear story for service delivery and operational support."),
                    new PowerPointCaseStudySection("Challenge",
                        "The client needed better device availability, consistent support, and less manual coordination across many locations."),
                    new PowerPointCaseStudySection("Solution",
                        "A structured service model combined standardized devices, support ownership, monitoring, and reporting."),
                    new PowerPointCaseStudySection("Result",
                        "The project summary keeps the executive story, metrics, and visual emphasis in one editable slide.")
                },
                new[] {
                    new PowerPointMetric("150", "devices"),
                    new PowerPointMetric("60", "locations")
                },
                "case-study",
                options => {
                    options.BrandText = "OFFICEIMO";
                    options.BandLabel = "Project portfolio";
                    options.Variant = PowerPointCaseStudyLayoutVariant.EditorialSplit;
                });

            deck.AddProcessSlide("How we work",
                "Transparent phases reduce risk and speed up delivery",
                new[] {
                    new PowerPointProcessStep("Analysis", "Understand the environment, needs, and business constraints."),
                    new PowerPointProcessStep("Discovery", "Review configuration, process maturity, and dependencies."),
                    new PowerPointProcessStep("Recommendations", "Prioritize actions and define the target architecture."),
                    new PowerPointProcessStep("Implementation", "Deliver changes in controlled stages."),
                    new PowerPointProcessStep("Care", "Keep the environment stable after rollout.")
                },
                "process",
                options => options.Variant = PowerPointProcessLayoutVariant.NumberedColumns);

            deck.AddCardGridSlide("Scope of services",
                "Reusable cards choose the placement grid for you.",
                new[] {
                    new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot", "Policy baseline" }),
                    new PowerPointCardContent("Maintenance", new[] { "Incident handling", "Monitoring", "Optimization" }),
                    new PowerPointCardContent("Consulting", new[] { "Roadmap", "Architecture", "Discovery" }),
                    new PowerPointCardContent("Audits", new[] { "Configuration", "Security review", "Modernization plan" })
                },
                "card-grid",
                options => {
                    options.SupportingText = "Use this when the deck needs strong structure but the content should remain editable.";
                    options.Variant = PowerPointCardGridLayoutVariant.SoftTiles;
                });

            deck.AddLogoWallSlide("Competence proof",
                "Logo and certification walls stay editable, but can still feel designed.",
                new[] {
                    new PowerPointLogoItem("Xerox", "Authorized service"),
                    new PowerPointLogoItem("Lenovo", "Partner"),
                    new PowerPointLogoItem("Brother", "Print"),
                    new PowerPointLogoItem("Samsung", "Devices"),
                    new PowerPointLogoItem("Epson", "Service"),
                    new PowerPointLogoItem("ASUS", "Hardware"),
                    new PowerPointLogoItem("Ricoh", "Print"),
                    new PowerPointLogoItem("Xiaomi", "Mobile")
                },
                "logo-wall",
                options => {
                    options.FeatureTitle = "Audit-ready proof area";
                    options.SupportingText = "Use real logo image paths when available; otherwise names remain editable.";
                });

            deck.AddCoverageSlide("Service coverage",
                "Normalized pins create a map-like layout without needing a custom map asset.",
                new[] {
                    new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
                    new PowerPointCoverageLocation("Szczecin", 0.18, 0.30),
                    new PowerPointCoverageLocation("Olsztyn", 0.67, 0.27),
                    new PowerPointCoverageLocation("Warszawa", 0.60, 0.48),
                    new PowerPointCoverageLocation("Poznan", 0.34, 0.46),
                    new PowerPointCoverageLocation("Wroclaw", 0.36, 0.70),
                    new PowerPointCoverageLocation("Krakow", 0.58, 0.78),
                    new PowerPointCoverageLocation("Rzeszow", 0.75, 0.76)
                },
                "coverage",
                options => {
                    options.SupportingText = "Regional teams and service locations";
                    options.MapLabel = "8 editable locations";
                });

            deck.AddCapabilitySlide("Service capability",
                "Structured text plus visual support for content-heavy service slides.",
                new[] {
                    new PowerPointCapabilitySection("Warranty and post-warranty service",
                        "Nationwide service operations for distributed environments.",
                        new[] { "Computers, notebooks, tablets", "Printers and scanners" }),
                    new PowerPointCapabilitySection("Extended care",
                        "Support beyond standard vendor warranty.",
                        new[] { "Individual SLA options", "Continuity-focused monitoring" }),
                    new PowerPointCapabilitySection("Operational gain",
                        "Keep the content readable without manual layout.",
                        new[] { "Clear ownership", "Consistent service story" })
                },
                "capability",
                options => {
                    options.VisualKind = PowerPointCapabilityVisualKind.CoverageMap;
                    options.VisualLabel = "Service locations";
                    options.Locations.Add(new PowerPointCoverageLocation("Warszawa", 0.60, 0.48));
                    options.Locations.Add(new PowerPointCoverageLocation("Gdansk", 0.55, 0.18));
                    options.Locations.Add(new PowerPointCoverageLocation("Wroclaw", 0.36, 0.70));
                    options.Locations.Add(new PowerPointCoverageLocation("Krakow", 0.58, 0.78));
                    options.Metrics.Add(new PowerPointMetric("12", "teams"));
                    options.Metrics.Add(new PowerPointMetric("8", "locations"));
                });

            deck.ComposeSlide(composer => {
                composer.AddTitle("Raw composition", "Use primitives when the slide needs its own structure.");
                PowerPointLayoutBox[] rows = composer.ContentRows(2, gutterCm: 0.55, topCm: 3.75);
                composer.AddCardGrid(new[] {
                    new PowerPointCardContent("Story", new[] { "Title block", "Narrative area" }),
                    new PowerPointCardContent("Evidence", new[] { "Metrics", "Visual support" }),
                    new PowerPointCardContent("Outcome", new[] { "Summary", "Next step" })
                }, rows[0], new PowerPointCardGridSlideOptions {
                    MaxColumns = 3,
                    Variant = PowerPointCardGridLayoutVariant.AccentTop
                });

                PowerPointLayoutBox[] lowerColumns = rows[1].SplitColumnsCm(2, 0.65);
                composer.AddCalloutBand("Composer regions keep the slide flexible without requiring manual coordinates.",
                    lowerColumns[0].TakeTopCm(1.45));
                composer.AddMetricStrip(new[] {
                    new PowerPointMetric("3", "primitives"),
                    new PowerPointMetric("1", "shared grid")
                }, lowerColumns[1].TakeTopCm(1.45));
            }, "composed", options => options.FooterRight = "Composable");

            presentation.Save();

            List<ValidationErrorInfo> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string details = string.Join(Environment.NewLine, errors.Take(5).Select(error => error.Description));
                throw new InvalidOperationException($"PowerPoint validation failed with {errors.Count} error(s).{Environment.NewLine}{details}");
            }

            Console.WriteLine($"    Saved: {filePath}");
            Console.WriteLine("    Validation: no Open XML errors found.");
            Helpers.Open(filePath, openPowerPoint);
        }
    }
}
