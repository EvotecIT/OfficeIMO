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

            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
            presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
            PowerPointDeckComposer deck = presentation.UseDesigner(brief, selected.Design.Index);

            deck.AddSectionSlide("Design brief recommendations",
                "Preview several directions, then choose one before rendering.",
                "brief-cover",
                options => options.SectionVariant = PowerPointSectionLayoutVariant.Poster);

            deck.AddCardGridSlide("Generated alternatives",
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

            deck.AddProcessSlide("Caller workflow",
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

            deck.ComposeSlide(composer => {
                composer.AddTitle("Selected direction", selected.Design.DirectionName);
                PowerPointLayoutBox[] columns = composer.ContentColumns(2, 0.8, topCm: 3.85);

                composer.AddCardGrid(selected.Reasons.Take(4).Select((reason, index) =>
                        new PowerPointCardContent("Reason " + (index + 1), new[] { reason })),
                    columns[0],
                    new PowerPointCardGridSlideOptions {
                        MaxColumns = 1,
                        Variant = PowerPointCardGridLayoutVariant.SoftTiles
                    });

                composer.AddMetricStrip(new[] {
                    new PowerPointMetric(selected.PreferenceScore.ToString(), "preference score"),
                    new PowerPointMetric((selected.Design.Index + 1).ToString(), "selected option"),
                    new PowerPointMetric(selected.Reasons.Count.ToString(), "reasons")
                }, columns[1].TakeTopCm(2.2));

                composer.AddCalloutBand(
                    $"{selected.Design.Mood} mood, {selected.Design.VisualStyle} visuals, " +
                    $"{selected.Design.HeadingFontName} headings. Change only the alternative index to render a different direction.",
                    columns[1].InsetCm(0, 2.75, 0, 0).TakeTopCm(1.75));
            }, "selected-direction", options => options.FooterRight = "Recommended");

            presentation.Save();
            Validate(filePath, presentation);
            Helpers.Open(filePath, openPowerPoint);
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
