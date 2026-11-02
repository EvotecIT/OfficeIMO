using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates choosing a generated design direction before rendering slides.
    /// </summary>
    public static class DesignBriefRecommendationsPowerPoint {
        public static void Example_DesignBriefRecommendationsPowerPoint(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Design brief recommendations");
            string filePath = Path.Join(folderPath, "PowerPoint Design Brief Recommendations.pptx");

            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "design-brief-recommendations", "technical rollout proposal")
                .WithIdentity("Client Theme", eyebrow: "OfficeIMO.PowerPoint", footerLeft: "OFFICEIMO",
                    footerRight: "Design brief")
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof)
                .WithPalette(surfaceColor: "#F6FAFC", panelBorderColor: "#D5E3EA");

            IReadOnlyList<PowerPointDeckDesignRecommendation> recommendations = brief.RecommendAlternatives(4);
            PowerPointDeckDesignRecommendation selected = recommendations
                .OrderByDescending(recommendation => recommendation.PreferenceScore)
                .ThenBy(recommendation => recommendation.Design.Index)
                .First();

            using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDeckPlan plan = new PowerPointDeckPlan();

            plan.AddSection("Design brief recommendations",
                "Preview several directions, then choose one before rendering.",
                "brief-cover",
                options => options.SectionVariant = PowerPointSectionLayoutVariant.Poster);

            plan.AddCardGrid("Generated alternatives",
                "The caller can inspect direction, mood, density, fonts, colors, and recommendation reasons.",
                recommendations.Select(recommendation => new PowerPointCardContent(
                    recommendation.Design.DirectionName,
                    new[] {
                        $"Score: {recommendation.PreferenceScore}",
                        $"{recommendation.Design.PaletteStyle} / {recommendation.Design.LayoutStrategy}",
                        $"{recommendation.Design.Mood} / {recommendation.Design.VisualStyle}",
                        $"{recommendation.Design.HeadingFontName} + {recommendation.Design.BodyFontName}"
                    },
                    recommendation.Design.AccentColor)),
                "brief-alternatives",
                options => {
                    options.Variant = PowerPointCardGridLayoutVariant.AccentTop;
                    options.MaxColumns = 4;
                    options.SupportingText = "Recommendations are explainable, not hidden template magic.";
                });

            plan.AddProcess("Caller workflow",
                "Keep the API simple while preserving room for different visual outcomes.",
                new[] {
                    new PowerPointProcessStep("Describe", "Create deterministic alternatives from brand, seed, and purpose."),
                    new PowerPointProcessStep("Recommend", "Rank alternatives against explicit design preferences."),
                    new PowerPointProcessStep("Select", "Use the chosen alternative index for the deck composer."),
                    new PowerPointProcessStep("Compose", "Mix semantic slides with raw composition primitives.")
                },
                "brief-workflow",
                options => {
                    options.Variant = PowerPointProcessLayoutVariant.Rail;
                    options.ConnectorStyle = PowerPointProcessConnectorStyle.SegmentArrows;
                });

            // The selected direction is ordinary content: reasons on the left, the numbers in a results panel.
            var direction = new List<PowerPointLayoutBlock>();
            direction.AddRange(selected.Reasons.Take(3).Select((reason, index) =>
                (PowerPointLayoutBlock)new PowerPointLayoutCard("Reason " + (index + 1), reason)));
            direction.Add(new PowerPointLayoutKpi(selected.PreferenceScore.ToString(), "Preference score"));
            direction.Add(new PowerPointLayoutKpi((selected.Design.Index + 1).ToString(), "Selected option"));
            direction.Add(new PowerPointLayoutKpi(selected.Reasons.Count.ToString(), "Reasons"));
            plan.Add(new PowerPointContentPlanSlide("Selected direction", direction, PowerPointLayout.CaseStudy, selected.Design.DirectionName) {
                Notes = $"{selected.Design.Mood} mood, {selected.Design.HeadingFontName} headings. Change only the alternative index to render a different direction."
            });

            PowerPointCompositionOptions composition = PowerPointCompositionOptions.FromBrief(brief);
            composition.SelectBestAlternative = false;
            composition.AlternativeIndex = selected.Design.Index;
            presentation.Compose(plan, composition);
            presentation.Save();
            Validate(filePath, presentation);
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }

        private static void Validate(string filePath, PowerPointDocument presentation) {
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
