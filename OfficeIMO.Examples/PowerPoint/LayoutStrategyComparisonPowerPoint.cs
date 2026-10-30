using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Renders one semantic plan in several built-in themes. The plan is unchanged; the theme supplies palette,
    /// fonts and type scale, and the slide layouts choose the geometry from the content.
    /// </summary>
    public static class LayoutStrategyComparisonPowerPoint {
        public static void Example_LayoutStrategyComparisonPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - One plan in three themes");
            PowerPointDeckPlan plan = CreatePlan();
            string? first = null;
            foreach ((string name, OfficeTheme theme) in new[] {
                         ("Modern", OfficeTheme.Modern), ("Boardroom", OfficeTheme.Boardroom), ("Dark", OfficeTheme.Dark)
                     }) {
                string filePath = Path.Join(folderPath, "PowerPoint Layout Strategy Comparison - " + name + ".pptx");
                first ??= filePath;
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath, new PowerPointCreateOptions { Theme = theme });
                // The design repeats the theme's own accent and fonts, so composition keeps the theme as it is.
                PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#" + theme.Palette.Accent1.ToRgbHex(), "layout-comparison-" + name,
                    headingFontName: theme.Fonts.Heading.Family, bodyFontName: theme.Fonts.Body.Family, footerLeft: "OfficeIMO · " + name + " theme");
                PowerPointCompositionOptions options = PowerPointCompositionOptions.FromDesign(design);
                options.ApplyTheme = false;
                presentation.Compose(plan, options);
                presentation.Save();
                Validate(filePath, presentation);
            }

            if (openPowerPoint && first != null) ExampleFileLauncher.Open(first);
        }

        private static PowerPointDeckPlan CreatePlan() {
            return new PowerPointDeckPlan()
                .AddSection("Service portfolio",
                    "One plan, three themes, no coordinates.",
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
                    "Five steps on one row; the layout follows the step count.",
                    new[] {
                        new PowerPointProcessStep("Discover", "Collect constraints and service expectations."),
                        new PowerPointProcessStep("Design", "Choose target architecture and rollout rules."),
                        new PowerPointProcessStep("Pilot", "Validate with a controlled group."),
                        new PowerPointProcessStep("Roll out", "Deliver in waves with reporting."),
                        new PowerPointProcessStep("Operate", "Move into repeatable support.")
                    },
                    "layout-process")
                .AddCardGrid("Scope of services",
                    "Cards stay semantic; the theme supplies their colour and type.",
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
            List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string details = string.Join(Environment.NewLine, errors.Take(5).Select(error => error.Description));
                throw new InvalidOperationException($"PowerPoint validation failed with {errors.Count} error(s).{Environment.NewLine}{details}");
            }

            Console.WriteLine($"    Saved: {filePath}");
            Console.WriteLine("    Validation: no Open XML errors found.");
        }
    }
}
