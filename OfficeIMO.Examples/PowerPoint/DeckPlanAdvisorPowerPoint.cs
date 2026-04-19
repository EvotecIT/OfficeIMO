using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates scoring a semantic deck plan before rendering it.
    /// </summary>
    public static class DeckPlanAdvisorPowerPoint {
        public static void Example_DeckPlanAdvisorPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Deck plan advisor");
            string filePath = Path.Join(folderPath, "PowerPoint Deck Plan Advisor.pptx");

            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#0B7FAB", "deck-plan-advisor", "service portfolio and implementation proposal")
                .WithIdentity("Service Portfolio", eyebrow: "OfficeIMO.PowerPoint", footerLeft: "OFFICEIMO",
                    footerRight: "Deck plan")
                .WithPalette(secondaryAccentColor: "#7C3AED", tertiaryAccentColor: "#14B8A6",
                    warmAccentColor: "#F59E0B", surfaceColor: "#F8FAFC", panelBorderColor: "#D7E2EA")
                .WithVariety(PowerPointDesignVariety.Exploratory)
                .WithPreferredMoods(PowerPointDesignMood.Editorial, PowerPointDesignMood.Energetic)
                .WithPreferredDensities(PowerPointSlideDensity.Balanced, PowerPointSlideDensity.Relaxed);

            PowerPointDeckPlan plan = CreatePlan();
            IReadOnlyList<PowerPointDeckPlanAlternativeSummary> alternatives =
                brief.DescribeDeckPlanAlternatives(plan, 4);
            PowerPointDeckPlanAlternativeSummary selected = alternatives
                .OrderByDescending(alternative => alternative.ContentFitScore)
                .ThenBy(alternative => alternative.Index)
                .First();

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDeckComposer deck = presentation.UseDesigner(brief, selected.Index);

            deck.AddSlides(plan);
            deck.ComposeSlide(composer => {
                composer.AddTitle("Why this alternative wins", selected.Design.DirectionName);
                PowerPointCompositionLayout layout = composer.UsePreset(PowerPointCompositionPreset.MetricStory,
                    PowerPointCompositionVariant.VisualLead);

                composer.AddCardGrid(selected.ContentFitReasons.Take(4).Select((reason, index) =>
                        new PowerPointCardContent("Fit signal " + (index + 1), new[] { reason })),
                    layout.Primary,
                    new PowerPointCardGridSlideOptions {
                        MaxColumns = 1,
                        Variant = PowerPointCardGridLayoutVariant.SoftTiles
                    });

                composer.AddVisualFrame(layout.Visual);
                composer.AddMetricStrip(new[] {
                    new PowerPointMetric(selected.ContentFitScore.ToString(), "fit score"),
                    new PowerPointMetric(selected.Slides.Count.ToString(), "planned slides"),
                    new PowerPointMetric(selected.Diagnostics.Count.ToString(), "diagnostics")
                }, layout.Metrics);
            }, "advisor-summary", options => options.FooterRight = "Fit score " + selected.ContentFitScore);

            presentation.Save();
            Validate(filePath, presentation);
            Helpers.Open(filePath, openPowerPoint);
        }

        private static PowerPointDeckPlan CreatePlan() {
            return new PowerPointDeckPlan()
                .AddSection("Service proposal",
                    "A semantic plan keeps the story reusable while the design can change.",
                    "advisor-cover",
                    options => options.SectionVariant = PowerPointSectionLayoutVariant.EditorialRail)
                .AddCaseStudy("Managed workplace rollout",
                    new[] {
                        new PowerPointCaseStudySection("Client",
                            "A distributed organization needed a clear service story for device rollout and support."),
                        new PowerPointCaseStudySection("Challenge",
                            "Many locations, mixed hardware, and manual coordination made delivery difficult to explain."),
                        new PowerPointCaseStudySection("Solution",
                            "Standardized onboarding, monitoring, and operating roles replaced ad hoc delivery."),
                        new PowerPointCaseStudySection("Result",
                            "The proposal can show outcomes, metrics, and visual emphasis without hand placement.")
                    },
                    new[] {
                        new PowerPointMetric("18", "sites"),
                        new PowerPointMetric("420", "devices"),
                        new PowerPointMetric("6", "workstreams")
                    },
                    "advisor-case-study")
                .AddProcess("Implementation path",
                    "The plan describes content intent; the chosen design handles layout.",
                    new[] {
                        new PowerPointProcessStep("Discover", "Collect constraints, dependencies, and service expectations."),
                        new PowerPointProcessStep("Design", "Choose target architecture and rollout rules."),
                        new PowerPointProcessStep("Pilot", "Validate the model with a controlled user group."),
                        new PowerPointProcessStep("Roll out", "Deliver in waves with clear reporting."),
                        new PowerPointProcessStep("Operate", "Move into repeatable support and optimization.")
                    },
                    "advisor-process")
                .AddCardGrid("Scope of services",
                    "Cards are semantic, so the same plan can choose a different grid.",
                    new[] {
                        new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot", "Policy baseline" }),
                        new PowerPointCardContent("Operations", new[] { "Incident handling", "Monitoring", "Optimization" }),
                        new PowerPointCardContent("Consulting", new[] { "Roadmap", "Architecture", "Discovery" }),
                        new PowerPointCardContent("Audits", new[] { "Configuration", "Security review", "Modernization plan" })
                    },
                    "advisor-cards")
                .AddCoverage("Delivery coverage",
                    "Normalized locations create an editable map-like view.",
                    new[] {
                        new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
                        new PowerPointCoverageLocation("Warszawa", 0.60, 0.48),
                        new PowerPointCoverageLocation("Poznan", 0.34, 0.46),
                        new PowerPointCoverageLocation("Wroclaw", 0.36, 0.70),
                        new PowerPointCoverageLocation("Krakow", 0.58, 0.78),
                        new PowerPointCoverageLocation("Rzeszow", 0.75, 0.76)
                    },
                    "advisor-coverage")
                .AddCapability("Operating model",
                    "Capability slides combine longer text with visual support.",
                    new[] {
                        new PowerPointCapabilitySection("Governance",
                            "Clear ownership for scope, risk, change, and service acceptance.",
                            new[] { "Decision log", "Change rules", "Service reporting" }),
                        new PowerPointCapabilitySection("Delivery",
                            "Repeatable rollout activities with fewer manual layout decisions.",
                            new[] { "Wave plan", "Readiness checks", "Pilot feedback" }),
                        new PowerPointCapabilitySection("Run",
                            "Transition from project delivery into stable operations.",
                            new[] { "SLA model", "Monitoring", "Continuous improvement" })
                    },
                    "advisor-capability");
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
    }
}
