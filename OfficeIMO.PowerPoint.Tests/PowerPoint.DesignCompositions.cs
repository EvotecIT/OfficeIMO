using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointDesignCompositions {
        [Fact]
        public void DesignerCompositions_CreateValidEditableDeck() {
            string filePath = CreateTempPresentationPath();

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                    presentation.ApplyDesignerTheme();

                    presentation.AddDesignerSectionSlide("Case Study", "Project portfolio",
                        options: new PowerPointDesignerSlideOptions {
                            Eyebrow = "OfficeIMO.PowerPoint",
                            FooterLeft = "OfficeIMO",
                            FooterRight = "The Good Slides"
                        });

                    presentation.AddDesignerProcessSlide("How we work",
                        "Clear phases reduce risk and speed up delivery",
                        new[] {
                            new PowerPointProcessStep("Discovery", "Map the environment and constraints."),
                            new PowerPointProcessStep("Design", "Select the strongest architecture and scope."),
                            new PowerPointProcessStep("Delivery", "Implement changes with controlled rollout."),
                            new PowerPointProcessStep("Care", "Keep the environment stable after go-live.")
                        });

                    presentation.AddDesignerCardGridSlide("Scope of services",
                        "Reusable cards pick their own grid.",
                        new[] {
                            new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot", "Policy baseline" }),
                            new PowerPointCardContent("Maintenance", new[] { "Incidents", "Monitoring", "Optimization" }),
                            new PowerPointCardContent("Consulting", new[] { "Roadmap", "Architecture", "Discovery" }),
                            new PowerPointCardContent("Audits", new[] { "Configuration", "Security", "Modernization" })
                        },
                        options: new PowerPointCardGridSlideOptions {
                            SupportingText = "Most common areas can be expressed as cards, tags, or metrics without manual coordinates."
                        });

                    PowerPointCaseStudySlideOptions caseOptions = new() {
                        Eyebrow = "OfficeIMO.PowerPoint",
                        FooterLeft = "OfficeIMO",
                        FooterRight = "The Good Slides",
                        BrandText = "OFFICEIMO",
                        BandLabel = "Project portfolio"
                    };
                    caseOptions.Tags.Add("Services");
                    caseOptions.Tags.Add("Monitoring");
                    caseOptions.Tags.Add("Systems");
                    caseOptions.Tags.Add("Case Study");

                    presentation.AddDesignerCaseStudySlide("Example client",
                        new[] {
                            new PowerPointCaseStudySection("Client", "A national organization needed a concise service story."),
                            new PowerPointCaseStudySection("Challenge", "The source material mixed details, outcomes, and implementation context."),
                            new PowerPointCaseStudySection("Solution", "A structured slide separates narrative columns, metrics, and visual support."),
                            new PowerPointCaseStudySection("Result", "The output remains editable and keeps visual hierarchy without hand-placement.")
                        },
                        new[] {
                            new PowerPointMetric("150", "devices"),
                            new PowerPointMetric("60", "locations")
                        },
                        options: caseOptions);

                    List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    Assert.Equal(4, presentation.Slides.Count);
                    Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes),
                        textBox => textBox.Text == "Scope of services");
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCardGrid_KeepsCardsWithinSlideBounds() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCardGridSlide("Services", null,
                    Enumerable.Range(1, 6).Select(index =>
                        new PowerPointCardContent("Card " + index, new[] { "One", "Two" })),
                    options: new PowerPointCardGridSlideOptions { MaxColumns = 3 });

                PowerPointLayoutBox slideBounds = new(0, 0, presentation.SlideSize.WidthEmus, presentation.SlideSize.HeightEmus);
                for (int i = 1; i <= 6; i++) {
                    PowerPointShape? card = slide.GetShape("Card " + i);
                    Assert.NotNull(card);
                    Assert.True(card!.Left >= slideBounds.Left, "Card is left of the slide.");
                    Assert.True(card.Top >= slideBounds.Top, "Card is above the slide.");
                    Assert.True(card.Right <= slideBounds.Right, "Card is right of the slide.");
                    Assert.True(card.Bottom <= slideBounds.Bottom, "Card is below the slide.");
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerProcessSlide_RejectsTooManyStepsForReadableLayout() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                IEnumerable<PowerPointProcessStep> steps = Enumerable.Range(1, 9)
                    .Select(index => new PowerPointProcessStep("Step " + index, "Body"));

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    presentation.AddDesignerProcessSlide("Too much", null, steps));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerSlideComposer_CompositionVariantsChangeRegionsWithoutLeavingContent() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointCompositionLayout? standard = null;
                PowerPointCompositionLayout? mirrored = null;
                PowerPointCompositionLayout? visualLead = null;
                PowerPointCompositionLayout? evidenceLead = null;
                var dashboardLayouts = new List<PowerPointCompositionLayout>();

                presentation.ComposeDesignerSlide(composer => {
                    standard = composer.UsePreset(PowerPointCompositionPreset.MetricStory,
                        PowerPointCompositionVariant.Standard);
                    mirrored = composer.UsePreset(PowerPointCompositionPreset.MetricStory,
                        PowerPointCompositionVariant.Mirrored);
                    visualLead = composer.UsePreset(PowerPointCompositionPreset.MetricStory,
                        PowerPointCompositionVariant.VisualLead);
                    evidenceLead = composer.UsePreset(PowerPointCompositionPreset.MetricStory,
                        PowerPointCompositionVariant.EvidenceLead);

                    foreach (PowerPointCompositionVariant variant in new[] {
                        PowerPointCompositionVariant.Standard,
                        PowerPointCompositionVariant.Mirrored,
                        PowerPointCompositionVariant.VisualLead,
                        PowerPointCompositionVariant.EvidenceLead
                    }) {
                        dashboardLayouts.Add(composer.UsePreset(PowerPointCompositionPreset.DashboardGrid, variant));
                    }
                });

                Assert.NotNull(standard);
                Assert.NotNull(mirrored);
                Assert.NotNull(visualLead);
                Assert.NotNull(evidenceLead);
                Assert.Equal(PowerPointCompositionVariant.Standard, standard!.Variant);
                Assert.Equal(PowerPointCompositionVariant.Mirrored, mirrored!.Variant);
                Assert.Equal(PowerPointCompositionVariant.VisualLead, visualLead!.Variant);
                Assert.Equal(PowerPointCompositionVariant.EvidenceLead, evidenceLead!.Variant);
                Assert.True(standard.Visual.LeftCm > standard.Primary.LeftCm);
                Assert.True(mirrored.Visual.LeftCm < mirrored.Primary.LeftCm);
                Assert.True(visualLead.Visual.WidthCm > standard.Visual.WidthCm);
                Assert.Equal(evidenceLead.Content.TopCm, evidenceLead.Metrics.TopCm);

                foreach (PowerPointCompositionLayout layout in new[] { standard, mirrored, visualLead, evidenceLead }) {
                    Assert.True(layout.Primary.Left >= layout.Content.Left);
                    Assert.True(layout.Primary.Right <= layout.Content.Right);
                    Assert.True(layout.Visual.Left >= layout.Content.Left);
                    Assert.True(layout.Visual.Right <= layout.Content.Right);
                    Assert.True(layout.Metrics.Top >= layout.Content.Top);
                    Assert.True(layout.Metrics.Bottom <= layout.Content.Bottom);
                }

                foreach (PowerPointCompositionLayout layout in dashboardLayouts) {
                    Assert.Equal(4, layout.Grid.Distinct().Count());
                    Assert.Contains(layout.Primary, layout.Grid);
                    Assert.Contains(layout.Secondary, layout.Grid);
                    Assert.Contains(layout.Visual, layout.Grid);
                    Assert.Contains(layout.Metrics, layout.Grid);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerTheme_UsesPresentationGradeFontsByDefault() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.ModernBlue;

            Assert.Equal("Poppins", theme.HeadingFontName);
            Assert.Equal("Lato", theme.BodyFontName);
        }

        [Fact]
        public void DesignerTheme_FromBrandCreatesDistinctPalette() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#7A3E9D", "Brand Theme",
                headingFontName: "Aptos Display", bodyFontName: "Aptos");

            Assert.Equal("Brand Theme", theme.Name);
            Assert.Equal("7A3E9D", theme.AccentColor);
            Assert.NotEqual(theme.AccentColor, theme.AccentDarkColor);
            Assert.NotEqual(theme.AccentColor, theme.AccentLightColor);
            Assert.Equal("Aptos Display", theme.HeadingFontName);
            Assert.Equal("Aptos", theme.BodyFontName);
        }

        [Fact]
        public void DesignerTheme_WithVariationKeepsBrandAccentButChangesSupportingPalette() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme");

            PowerPointDesignTheme first = theme.WithVariation("client-a");
            PowerPointDesignTheme second = theme.WithVariation("client-b");

            Assert.Equal(theme.AccentColor, first.AccentColor);
            Assert.Equal(theme.AccentColor, second.AccentColor);
            Assert.NotEqual(first.Accent2Color, second.Accent2Color);
            Assert.NotEqual(first.Accent3Color, second.Accent3Color);
            Assert.NotEqual(first.WarningColor, second.WarningColor);
            Assert.Contains(first.WarningColor, new[] { "F4A100", "F59E0B", "E9B44C", "FF8A3D", "D4AF37", "F97316", "EF4444" });
        }

        [Fact]
        public void DesignerTheme_CanApplyNamedPaletteStyle() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme");

            PowerPointDesignTheme styled =
                theme.WithPaletteStyle(PowerPointPaletteStyle.SplitComplementary, "client-a");

            Assert.Equal("008C95", styled.AccentColor);
            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, styled.PaletteStyle);
            Assert.NotEqual(styled.AccentColor, styled.Accent2Color);
            Assert.NotEqual(styled.Accent2Color, styled.Accent3Color);
            Assert.NotEqual(theme.SurfaceColor, styled.SurfaceColor);
        }

        [Fact]
        public void DesignerTheme_AutoPaletteStyleIsDeterministic() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme");

            PowerPointDesignTheme first = theme.WithPaletteStyle(PowerPointPaletteStyle.Auto, "client-a");
            PowerPointDesignTheme second = theme.WithPaletteStyle(PowerPointPaletteStyle.Auto, "client-a");

            Assert.Equal(first.PaletteStyle, second.PaletteStyle);
            Assert.NotEqual(PowerPointPaletteStyle.Auto, first.PaletteStyle);
            Assert.Equal(first.Accent2Color, second.Accent2Color);
            Assert.Equal(first.Accent3Color, second.Accent3Color);
        }

        [Fact]
        public void DesignerTheme_CanApplyTypographyStyle() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme");

            PowerPointDesignTheme styled =
                theme.WithTypographyStyle(PowerPointTypographyStyle.EditorialSerif, "client-a");

            Assert.Equal(PowerPointTypographyStyle.EditorialSerif, styled.TypographyStyle);
            Assert.Equal("Georgia", styled.HeadingFontName);
            Assert.Equal("Aptos", styled.BodyFontName);
        }

        [Fact]
        public void DesignerTheme_EnergeticMoodMatchesFriendlySansFonts() {
            PowerPointDesignTheme energetic = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme")
                .WithMood(PowerPointDesignMood.Energetic);

            Assert.Equal(PowerPointTypographyStyle.FriendlySans, energetic.TypographyStyle);
            Assert.Equal("Poppins", energetic.HeadingFontName);
            Assert.Equal("Lato", energetic.BodyFontName);
        }

        [Fact]
        public void DesignerTheme_AutoTypographyStyleIsDeterministic() {
            PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("#008C95", "Brand Theme");

            PowerPointDesignTheme first = theme.WithTypographyStyle(PowerPointTypographyStyle.Auto, "client-a");
            PowerPointDesignTheme second = theme.WithTypographyStyle(PowerPointTypographyStyle.Auto, "client-a");

            Assert.Equal(first.TypographyStyle, second.TypographyStyle);
            Assert.NotEqual(PowerPointTypographyStyle.Auto, first.TypographyStyle);
            Assert.Equal(first.HeadingFontName, second.HeadingFontName);
            Assert.Equal(first.BodyFontName, second.BodyFontName);
        }

        [Fact]
        public void DesignerTheme_DerivedVariantsResetPaletteStyleMetadata() {
            PowerPointDesignTheme theme = PowerPointDesignTheme
                .FromBrand("#008C95", "Brand Theme")
                .WithPaletteStyle(PowerPointPaletteStyle.SplitComplementary, "client-a");

            PowerPointDesignTheme variation = theme.WithVariation("client-b");
            PowerPointDesignTheme corporate = theme.WithMood(PowerPointDesignMood.Corporate);
            PowerPointDesignTheme editorial = theme.WithMood(PowerPointDesignMood.Editorial);
            PowerPointDesignTheme energetic = theme.WithMood(PowerPointDesignMood.Energetic);
            PowerPointDesignTheme mood = theme.WithMood(PowerPointDesignMood.Minimal);

            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, theme.PaletteStyle);
            Assert.Equal(PowerPointPaletteStyle.Auto, variation.PaletteStyle);
            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, corporate.PaletteStyle);
            Assert.Equal(PowerPointPaletteStyle.Auto, editorial.PaletteStyle);
            Assert.Equal(PowerPointPaletteStyle.Auto, energetic.PaletteStyle);
            Assert.Equal(PowerPointPaletteStyle.Auto, mood.PaletteStyle);
            Assert.Equal(theme.AccentColor, variation.AccentColor);
            Assert.Equal(theme.Accent2Color, corporate.Accent2Color);
            Assert.Equal(theme.AccentColor, mood.AccentColor);
        }

        [Fact]
        public void DesignerDeckDesign_ConfiguresThemeIntentAndChromeFromOnePlace() {
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "client-a",
                PowerPointDesignMood.Editorial, name: "Client A", eyebrow: "Portfolio",
                footerLeft: "CLIENT A", footerRight: "Service deck");

            Assert.Equal("008C95", design.Theme.AccentColor);
            Assert.Equal(PowerPointDesignMood.Editorial, design.BaseIntent.Mood);
            Assert.Equal(PowerPointSlideDensity.Relaxed, design.BaseIntent.Density);
            Assert.Equal(PowerPointVisualStyle.Soft, design.BaseIntent.VisualStyle);
            Assert.Equal("Aptos Display", design.Theme.HeadingFontName);

            PowerPointCaseStudySlideOptions options = design.Configure(new PowerPointCaseStudySlideOptions {
                Variant = PowerPointCaseStudyLayoutVariant.VisualHero
            }, "case-study");

            Assert.Equal("Portfolio", options.Eyebrow);
            Assert.Equal("CLIENT A", options.FooterLeft);
            Assert.Equal("Service deck", options.FooterRight);
            Assert.Equal("client-a/case-study", options.DesignIntent.Seed);
            Assert.Equal(PowerPointDesignMood.Editorial, options.DesignIntent.Mood);
            Assert.Equal(PowerPointCaseStudyLayoutVariant.VisualHero, options.Variant);
        }

        [Fact]
        public void DesignerDeckDesign_CanCreateDistinctAlternativesFromOneBrand() {
            IReadOnlyList<PowerPointDeckDesign> alternatives =
                PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "client-a",
                    count: 5, name: "Client A", eyebrow: "Portfolio",
                    footerLeft: "CLIENT A", footerRight: "Service deck");

            Assert.Equal(5, alternatives.Count);
            Assert.All(alternatives, design => Assert.Equal("008C95", design.Theme.AccentColor));
            Assert.Equal(PowerPointDesignMood.Corporate, alternatives[0].BaseIntent.Mood);
            Assert.Equal(PowerPointDesignMood.Editorial, alternatives[1].BaseIntent.Mood);
            Assert.Equal(PowerPointDesignMood.Minimal, alternatives[2].BaseIntent.Mood);
            Assert.Equal(PowerPointDesignMood.Energetic, alternatives[3].BaseIntent.Mood);
            Assert.Equal(PowerPointVisualStyle.Soft, alternatives[4].BaseIntent.VisualStyle);
            Assert.Equal("Executive", alternatives[4].Direction.Name);
            Assert.Equal("Segoe UI Semibold", alternatives[4].Theme.HeadingFontName);
            Assert.NotEqual(alternatives[0].Theme.Accent2Color, alternatives[1].Theme.Accent2Color);
            Assert.False(alternatives[2].ShowDirectionMotif);
            Assert.Equal("Portfolio", alternatives[3].Options("cover").Eyebrow);
            Assert.Equal("client-a/direction-4/cover", alternatives[3].Options("cover").DesignIntent.Seed);
        }

        [Fact]
        public void DesignerDeckDesign_CanUseNamedCreativeDirection() {
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "executive-client",
                PowerPointDesignDirection.Executive, name: "Executive Client", footerLeft: "CLIENT");

            Assert.Equal("Executive", design.Direction.Name);
            Assert.Equal(PowerPointDesignMood.Corporate, design.BaseIntent.Mood);
            Assert.Equal(PowerPointVisualStyle.Soft, design.BaseIntent.VisualStyle);
            Assert.Equal("Segoe UI Semibold", design.Theme.HeadingFontName);
            Assert.Equal("Segoe UI", design.Theme.BodyFontName);
            Assert.False(design.ShowDirectionMotif);
            Assert.Equal("CLIENT", design.Options("cover").FooterLeft);
        }

        [Fact]
        public void DesignerDeckDesign_SummaryUsesActualTypographyMetadataForCustomDirectionFonts() {
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "signal-client",
                PowerPointDesignDirection.Signal, name: "Signal Client");

            PowerPointDeckDesignSummary summary = design.Describe();

            Assert.Equal("Poppins", summary.HeadingFontName);
            Assert.Equal("Aptos", summary.BodyFontName);
            Assert.Equal(PowerPointTypographyStyle.Auto, summary.TypographyStyle);
        }

        [Fact]
        public void DesignerDeckDesign_CanCreateAlternativesFromCustomDirections() {
            PowerPointDesignDirection boardBrief = new("Board Brief", PowerPointDesignMood.Corporate,
                PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
                showDirectionMotif: false);
            PowerPointDesignDirection fieldOps = new("Field Ops", PowerPointDesignMood.Energetic,
                PowerPointSlideDensity.Compact, PowerPointVisualStyle.Geometric, "Poppins", "Segoe UI");

            IReadOnlyList<PowerPointDeckDesign> alternatives =
                PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "custom-client",
                    new[] { boardBrief, fieldOps }, name: "Client", footerLeft: "CLIENT");

            Assert.Equal(2, alternatives.Count);
            Assert.Equal("Board Brief", alternatives[0].Direction.Name);
            Assert.Equal(PowerPointVisualStyle.Soft, alternatives[0].BaseIntent.VisualStyle);
            Assert.Equal("Georgia", alternatives[0].Theme.HeadingFontName);
            Assert.False(alternatives[0].ShowDirectionMotif);
            Assert.Equal("custom-client/board-brief-1/cover", alternatives[0].Options("cover").DesignIntent.Seed);

            Assert.Equal("Field Ops", alternatives[1].Direction.Name);
            Assert.Equal(PowerPointSlideDensity.Compact, alternatives[1].BaseIntent.Density);
            Assert.Equal("Segoe UI", alternatives[1].Theme.BodyFontName);
            Assert.True(alternatives[1].ShowDirectionMotif);
            Assert.Equal("CLIENT", alternatives[1].Options("cover").FooterLeft);
        }

        [Fact]
        public void DesignerDeckDesign_CanCreateAlternativesFromRecipe() {
            IReadOnlyList<PowerPointDeckDesign> alternatives =
                PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "service-client",
                    PowerPointDesignRecipe.ConsultingPortfolio, name: "Client", footerLeft: "CLIENT");

            Assert.Equal(3, alternatives.Count);
            Assert.Equal("Consulting Portfolio", PowerPointDesignRecipe.ConsultingPortfolio.Name);
            Assert.Equal("Board Story", alternatives[0].Direction.Name);
            Assert.Equal(PowerPointDesignMood.Corporate, alternatives[0].BaseIntent.Mood);
            Assert.Equal(PowerPointVisualStyle.Soft, alternatives[0].BaseIntent.VisualStyle);
            Assert.Equal("Georgia", alternatives[0].Theme.HeadingFontName);
            Assert.False(alternatives[0].ShowDirectionMotif);
            Assert.Equal("Project portfolio", alternatives[0].Options("cover").Eyebrow);
            Assert.Equal("service-client/consulting-portfolio-board-story-1/cover",
                alternatives[0].Options("cover").DesignIntent.Seed);

            Assert.Equal("Field Proof", alternatives[1].Direction.Name);
            Assert.Equal(PowerPointDesignMood.Energetic, alternatives[1].BaseIntent.Mood);
            Assert.Equal(PowerPointVisualStyle.Geometric, alternatives[1].BaseIntent.VisualStyle);
            Assert.True(alternatives[1].ShowDirectionMotif);

            Assert.Equal("Quiet Appendix", alternatives[2].Direction.Name);
            Assert.Equal(PowerPointDesignMood.Minimal, alternatives[2].BaseIntent.Mood);
            Assert.Equal(PowerPointVisualStyle.Minimal, alternatives[2].BaseIntent.VisualStyle);
            Assert.Equal("CLIENT", alternatives[2].Options("cover").FooterLeft);
        }

        [Fact]
        public void DesignerDeckDesign_RecipeAlternativesCanCycleAndOverrideFonts() {
            IReadOnlyList<PowerPointDeckDesign> alternatives =
                PowerPointDeckDesign.CreateAlternativesFromBrand("#008C95", "technical-client",
                    PowerPointDesignRecipe.TechnicalProposal, count: 4, name: "Client",
                    headingFontName: "Aptos Display", bodyFontName: "Aptos");

            Assert.Equal(4, alternatives.Count);
            Assert.Equal("Architecture Map", alternatives[0].Direction.Name);
            Assert.Equal("Architecture Map", alternatives[3].Direction.Name);
            Assert.NotEqual(alternatives[0].Options("cover").DesignIntent.Seed,
                alternatives[3].Options("cover").DesignIntent.Seed);
            Assert.All(alternatives, design => Assert.Equal("Aptos Display", design.Theme.HeadingFontName));
            Assert.All(alternatives, design => Assert.Equal("Aptos", design.Theme.BodyFontName));
            Assert.Equal("Technical proposal", alternatives[1].Options("cover").Eyebrow);
        }

        [Fact]
        public void DesignerRecipe_CanCreateAlternativesDirectlyAndMatchPurpose() {
            PowerPointDesignRecipe? recipe = PowerPointDesignRecipe.FindBuiltIn("board decision brief");

            Assert.Same(PowerPointDesignRecipe.ExecutiveBrief, recipe);

            IReadOnlyList<PowerPointDeckDesign> alternatives = recipe!.CreateAlternativesFromBrand("#1F6FEB",
                "board-pack", count: 2, name: "Board Pack", footerLeft: "BOARD");

            Assert.Equal(2, alternatives.Count);
            Assert.Equal("Decision Pack", alternatives[0].Direction.Name);
            Assert.Equal("Investment Memo", alternatives[1].Direction.Name);
            Assert.Equal("Executive summary", alternatives[0].Options("cover").Eyebrow);
            Assert.Equal("BOARD", alternatives[1].Options("cover").FooterLeft);
            Assert.Equal("board-pack/executive-brief-investment-memo-2/cover",
                alternatives[1].Options("cover").DesignIntent.Seed);
        }

        [Fact]
        public void DesignerRecipe_CanMatchTransformationRoadmaps() {
            PowerPointDesignRecipe? recipe = PowerPointDesignRecipe.FindBuiltIn("transformation roadmap");

            Assert.Same(PowerPointDesignRecipe.TransformationRoadmap, recipe);

            IReadOnlyList<PowerPointDeckDesign> alternatives = recipe!.CreateAlternativesFromBrand("#008C95",
                "roadmap-client", count: 3, name: "Roadmap Client", footerLeft: "ROADMAP");

            Assert.Equal(3, alternatives.Count);
            Assert.Equal("North Star", alternatives[0].Direction.Name);
            Assert.Equal(PowerPointDesignMood.Editorial, alternatives[0].BaseIntent.Mood);
            Assert.Equal("Georgia", alternatives[0].Theme.HeadingFontName);
            Assert.False(alternatives[0].ShowDirectionMotif);

            Assert.Equal("Momentum Map", alternatives[1].Direction.Name);
            Assert.Equal(PowerPointVisualStyle.Geometric, alternatives[1].BaseIntent.VisualStyle);
            Assert.True(alternatives[1].ShowDirectionMotif);

            Assert.Equal("Operating Plan", alternatives[2].Direction.Name);
            Assert.Equal(PowerPointSlideDensity.Compact, alternatives[2].BaseIntent.Density);
            Assert.Equal("Roadmap", alternatives[2].Options("cover").Eyebrow);
            Assert.Equal("ROADMAP", alternatives[2].Options("cover").FooterLeft);
            Assert.Equal("roadmap-client/transformation-roadmap-operating-plan-3/cover",
                alternatives[2].Options("cover").DesignIntent.Seed);
        }

        [Fact]
        public void DesignerRecipe_CanDescribeBuiltInsAndMatches() {
            IReadOnlyList<PowerPointDesignRecipeSummary> recipes = PowerPointDesignRecipe.DescribeBuiltIns();
            IReadOnlyList<PowerPointDesignRecipeSummary> matches =
                PowerPointDesignRecipe.DescribeMatches("roadmap program");

            Assert.Equal(4, recipes.Count);
            Assert.Equal("Consulting Portfolio", recipes[0].Name);
            Assert.Equal("Transformation Roadmap", recipes[3].Name);
            Assert.Equal("Roadmap", recipes[3].DefaultEyebrow);
            Assert.Contains("journey", recipes[3].Keywords);
            Assert.Equal(3, recipes[3].DirectionCount);
            Assert.Equal("Momentum Map", recipes[3].Directions[1].Name);
            Assert.Equal(PowerPointDesignMood.Energetic, recipes[3].Directions[1].Mood);
            Assert.Equal("Poppins", recipes[3].Directions[1].HeadingFontName);
            Assert.Contains("Transformation Roadmap", recipes[3].ToString());
            Assert.Contains("Momentum Map", recipes[3].Directions[1].ToString());

            PowerPointDesignRecipeSummary match = Assert.Single(matches);
            Assert.Equal("Transformation Roadmap", match.Name);
            Assert.Empty(PowerPointDesignRecipe.DescribeMatches(""));
        }

        [Fact]
        public void DesignerDesignBrief_CanDescribeCreativeDirectionPacks() {
            IReadOnlyList<PowerPointCreativeDirectionPackSummary> packs =
                PowerPointDesignBrief.DescribeCreativeDirectionPacks();
            PowerPointCreativeDirectionPackSummary fieldProof =
                PowerPointDesignBrief.DescribeCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof);
            PowerPointCreativeDirectionPackSummary auto =
                PowerPointDesignBrief.DescribeCreativeDirectionPack(PowerPointCreativeDirectionPack.Auto);

            Assert.Equal(5, packs.Count);
            Assert.Equal(0, packs[0].Index);
            Assert.Equal(PowerPointCreativeDirectionPack.Boardroom, packs[0].Pack);
            Assert.Equal("Boardroom", packs[0].Name);
            Assert.Equal("Executive Brief", packs[0].RecipeName);
            Assert.Equal(PowerPointPaletteStyle.CoolNeutral, packs[0].PaletteStyle);
            Assert.Equal(PowerPointAutoLayoutStrategy.ContentFirst, packs[0].LayoutStrategy);
            Assert.Equal(PowerPointDesignVariety.Balanced, packs[0].Variety);
            Assert.Contains(PowerPointDesignMood.Corporate, packs[0].PreferredMoods);
            Assert.Contains("Boardroom", packs[0].ToString());

            Assert.Equal("Field Proof", fieldProof.Name);
            Assert.Equal(packs[1].Index, fieldProof.Index);
            Assert.Equal("Consulting Portfolio", fieldProof.RecipeName);
            Assert.Equal(PowerPointDesignVariety.Exploratory, fieldProof.Variety);
            Assert.Contains(PowerPointVisualStyle.Geometric, fieldProof.PreferredVisualStyles);

            Assert.Equal(PowerPointCreativeDirectionPack.Auto, auto.Pack);
            Assert.Equal(-1, auto.Index);
            Assert.Null(auto.RecipeName);
            Assert.Null(auto.PaletteStyle);
            Assert.Empty(auto.PreferredMoods);
            Assert.Contains("Purpose matched", auto.ToString());
        }

        [Fact]
        public void DesignerDesignBrief_CreatesAlternativesFromPurposeAndIdentity() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-client", "technical rollout proposal")
                .WithIdentity("Client", footerLeft: "CLIENT")
                .WithFonts(bodyFontName: "Segoe UI");

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(2);

            Assert.Equal(2, alternatives.Count);
            Assert.Equal("Architecture Map", alternatives[0].Direction.Name);
            Assert.Equal("Runbook", alternatives[1].Direction.Name);
            Assert.Equal("Technical proposal", alternatives[0].Options("cover").Eyebrow);
            Assert.Equal("CLIENT", alternatives[1].Options("cover").FooterLeft);
            Assert.Equal("Segoe UI", alternatives[0].Theme.BodyFontName);
            Assert.Equal("Segoe UI", alternatives[1].Theme.BodyFontName);
            Assert.NotEqual(alternatives[0].Theme.Accent2Color, alternatives[1].Theme.Accent2Color);
        }

        [Fact]
        public void DesignerDesignBrief_CanRankRecipeDirectionsByPreferences() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-preferred", "technical rollout proposal")
                .WithPreferredMoods(PowerPointDesignMood.Energetic)
                .WithPreferredVisualStyles(PowerPointVisualStyle.Geometric);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(3);
            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(1);
            IReadOnlyList<PowerPointDeckDesignRecommendation> recommendations =
                brief.RecommendAlternatives(2);

            Assert.Equal("Delivery Signal", alternatives[0].Direction.Name);
            Assert.Equal("Architecture Map", alternatives[1].Direction.Name);
            Assert.Equal("Runbook", alternatives[2].Direction.Name);
            Assert.Equal("Technical proposal", alternatives[0].Options("cover").Eyebrow);
            Assert.Equal(PowerPointDesignMood.Energetic, summaries[0].Mood);
            Assert.Equal(PowerPointVisualStyle.Geometric, summaries[0].VisualStyle);
            Assert.Equal(PowerPointDesignMood.Energetic, brief.PreferredMoods[0]);
            Assert.Equal(PowerPointVisualStyle.Geometric, brief.PreferredVisualStyles[0]);
            Assert.Equal("Delivery Signal", recommendations[0].Design.DirectionName);
            Assert.True(recommendations[0].MatchesPreferences);
            Assert.Equal(5, recommendations[0].PreferenceScore);
            Assert.Contains("Matches preferred mood: Energetic.", recommendations[0].Reasons);
            Assert.Contains("Matches preferred visual style: Geometric.", recommendations[0].Reasons);
            Assert.Contains("Delivery Signal score 5", recommendations[0].ToString());
            Assert.True(recommendations[1].MatchesPreferences);
            Assert.Equal(2, recommendations[1].PreferenceScore);

            brief.ClearDesignPreferences();
            Assert.Empty(brief.PreferredMoods);
            Assert.Equal("Architecture Map", brief.CreateAlternatives(1)[0].Direction.Name);
        }

        [Fact]
        public void DesignerDesignBrief_CanControlAlternativeVariety() {
            PowerPointDesignBrief focused = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-focused", "technical rollout proposal")
                .WithPreferredMoods(PowerPointDesignMood.Energetic)
                .WithVariety(PowerPointDesignVariety.Focused);
            PowerPointDesignBrief exploratory = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-exploratory", "technical rollout proposal")
                .WithVariety(PowerPointDesignVariety.Exploratory);

            IReadOnlyList<PowerPointDeckDesign> focusedAlternatives = focused.CreateAlternatives(3);
            IReadOnlyList<PowerPointDeckDesignSummary> exploratorySummaries = exploratory.DescribeAlternatives(5);

            Assert.Equal(PowerPointDesignVariety.Focused, focused.Variety);
            Assert.All(focusedAlternatives,
                design => Assert.Equal("Delivery Signal", design.Direction.Name));
            Assert.NotEqual(focusedAlternatives[0].Seed, focusedAlternatives[1].Seed);
            Assert.Equal("Architecture Map", exploratorySummaries[0].DirectionName);
            Assert.Equal("Runbook", exploratorySummaries[1].DirectionName);
            Assert.Equal("Delivery Signal", exploratorySummaries[2].DirectionName);
            Assert.Equal("Structured", exploratorySummaries[3].DirectionName);
            Assert.Equal("Editorial", exploratorySummaries[4].DirectionName);
        }

        [Fact]
        public void DesignerDesignBrief_CanDescribeAlternativesBeforeChoosing() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-preview", "technical rollout proposal")
                .WithIdentity("Client")
                .WithFonts(bodyFontName: "Segoe UI");

            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(3);

            Assert.Equal(3, summaries.Count);
            Assert.Equal(0, summaries[0].Index);
            Assert.Equal("Architecture Map", summaries[0].DirectionName);
            Assert.Equal("Runbook", summaries[1].DirectionName);
            Assert.Equal(PowerPointDesignMood.Minimal, summaries[1].Mood);
            Assert.Equal("008C95", summaries[2].AccentColor);
            Assert.Equal("Segoe UI", summaries[2].BodyFontName);
            Assert.NotEqual(summaries[0].Accent2Color, summaries[1].Accent2Color);
            Assert.Contains("Delivery Signal", summaries[2].ToString());
        }

        [Fact]
        public void DesignerDesignBrief_CanOverrideSupportingPaletteWithoutCustomTheme() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-palette", "technical rollout proposal")
                .WithPalette(
                    secondaryAccentColor: "#6D5BD0",
                    tertiaryAccentColor: "24A148",
                    warmAccentColor: "#FFB000",
                    surfaceColor: "F2F6F8",
                    panelBorderColor: "#B8C7D3");

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(2);
            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(1);

            Assert.Equal("008C95", alternatives[0].Theme.AccentColor);
            Assert.Equal("6D5BD0", alternatives[0].Theme.Accent2Color);
            Assert.Equal("24A148", alternatives[0].Theme.Accent3Color);
            Assert.Equal("FFB000", alternatives[0].Theme.WarningColor);
            Assert.Equal("F2F6F8", alternatives[0].Theme.SurfaceColor);
            Assert.Equal("B8C7D3", alternatives[0].Theme.PanelBorderColor);
            Assert.Equal("6D5BD0", alternatives[1].Theme.Accent2Color);
            Assert.Equal("6D5BD0", summaries[0].Accent2Color);
            Assert.Equal("FFB000", summaries[0].WarningColor);
        }

        [Fact]
        public void DesignerDesignBrief_CanChoosePaletteStyleBeforeManualOverrides() {
            PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#008C95", "client-demo")
                .WithPaletteStyle(PowerPointPaletteStyle.Monochrome)
                .WithPalette(secondaryAccentColor: "#6D5BD0");

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(2);
            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(1);

            Assert.Equal(PowerPointPaletteStyle.Monochrome, brief.PaletteStyle);
            Assert.All(alternatives, design => Assert.Equal(PowerPointPaletteStyle.Monochrome,
                design.Theme.PaletteStyle));
            Assert.All(alternatives, design => Assert.Equal("6D5BD0", design.Theme.Accent2Color));
            Assert.Equal(PowerPointPaletteStyle.Monochrome, summaries[0].PaletteStyle);
            Assert.Equal("6D5BD0", summaries[0].Accent2Color);
        }

        [Fact]
        public void DesignerDesignBrief_CanChooseTypographyStyleBeforeManualFontOverrides() {
            PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#008C95", "client-demo")
                .WithTypographyStyle(PowerPointTypographyStyle.EditorialSerif)
                .WithFonts(bodyFontName: "Segoe UI");

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(2);
            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(1);

            Assert.Equal(PowerPointTypographyStyle.EditorialSerif, brief.TypographyStyle);
            Assert.All(alternatives, design => Assert.Equal(PowerPointTypographyStyle.EditorialSerif,
                design.Theme.TypographyStyle));
            Assert.All(alternatives, design => Assert.Equal("Georgia", design.Theme.HeadingFontName));
            Assert.All(alternatives, design => Assert.Equal("Segoe UI", design.Theme.BodyFontName));
            Assert.Equal(PowerPointTypographyStyle.EditorialSerif, summaries[0].TypographyStyle);
            Assert.Equal("Segoe UI", summaries[0].BodyFontName);
        }

        [Fact]
        public void DesignerDesignBrief_CanApplyCreativeDirectionPack() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-pack", "technical rollout proposal")
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(3);
            IReadOnlyList<PowerPointDeckDesignSummary> summaries = brief.DescribeAlternatives(1);
            IReadOnlyList<PowerPointDeckDesignRecommendation> recommendations =
                brief.RecommendAlternatives(2);

            Assert.Equal(PowerPointCreativeDirectionPack.FieldProof, brief.CreativeDirectionPack);
            Assert.Same(PowerPointDesignRecipe.ConsultingPortfolio, brief.Recipe);
            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, brief.PaletteStyle);
            Assert.Equal(PowerPointAutoLayoutStrategy.VisualFirst, brief.LayoutStrategy);
            Assert.Equal(PowerPointDesignVariety.Exploratory, brief.Variety);
            Assert.Equal("Field Proof", alternatives[0].Direction.Name);
            Assert.Equal("Board Story", alternatives[1].Direction.Name);
            Assert.Equal(PowerPointAutoLayoutStrategy.VisualFirst, alternatives[0].BaseIntent.LayoutStrategy);
            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, alternatives[0].Theme.PaletteStyle);
            Assert.Equal(PowerPointDesignMood.Energetic, summaries[0].Mood);
            Assert.Equal(PowerPointVisualStyle.Geometric, summaries[0].VisualStyle);
            Assert.True(recommendations[0].MatchesPreferences);
            Assert.Contains("Matches preferred mood: Energetic.", recommendations[0].Reasons);
        }

        [Fact]
        public void DesignerDesignBrief_ApplyingCreativeDirectionPackClearsStaleManualThemeOverrides() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-pack-stale", "technical rollout proposal")
                .WithPalette(secondaryAccentColor: "#6D5BD0", surfaceColor: "#F2F6F8")
                .WithFonts("Segoe UI Semibold", "Segoe UI")
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(1);

            Assert.Null(brief.SecondaryAccentColor);
            Assert.Null(brief.SurfaceColor);
            Assert.Null(brief.HeadingFontName);
            Assert.Null(brief.BodyFontName);
            Assert.Equal(PowerPointPaletteStyle.SplitComplementary, alternatives[0].Theme.PaletteStyle);
            Assert.NotEqual("6D5BD0", alternatives[0].Theme.Accent2Color);
            Assert.NotEqual("F2F6F8", alternatives[0].Theme.SurfaceColor);
            Assert.NotEqual("Segoe UI Semibold", alternatives[0].Theme.HeadingFontName);
            Assert.NotEqual("Segoe UI", alternatives[0].Theme.BodyFontName);
        }

        [Fact]
        public void DesignerDesignBrief_AutoCreativeDirectionPackClearsPackOverrides() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-pack-auto", "technical rollout proposal")
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof)
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.Auto);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(1);

            Assert.Equal(PowerPointCreativeDirectionPack.Auto, brief.CreativeDirectionPack);
            Assert.Null(brief.Recipe);
            Assert.Null(brief.PaletteStyle);
            Assert.Null(brief.TypographyStyle);
            Assert.Null(brief.LayoutStrategy);
            Assert.Equal(PowerPointDesignVariety.Balanced, brief.Variety);
            Assert.Empty(brief.PreferredMoods);
            Assert.Empty(brief.PreferredDensities);
            Assert.Empty(brief.PreferredVisualStyles);
            Assert.Equal("Architecture Map", alternatives[0].Direction.Name);
            Assert.Equal(PowerPointAutoLayoutStrategy.ContentFirst, alternatives[0].BaseIntent.LayoutStrategy);
            Assert.Equal(PowerPointPaletteStyle.Auto, alternatives[0].Theme.PaletteStyle);
        }

        [Fact]
        public void DesignerDesignBrief_AutoCreativeDirectionPackPreservesExplicitTypography() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-pack-typography", "technical rollout proposal")
                .WithTypographyStyle(PowerPointTypographyStyle.EditorialSerif)
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.FieldProof)
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.Auto);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(1);

            Assert.Equal(PowerPointTypographyStyle.EditorialSerif, brief.TypographyStyle);
            Assert.Equal(PowerPointTypographyStyle.EditorialSerif, alternatives[0].Theme.TypographyStyle);
            Assert.Equal("Georgia", alternatives[0].Theme.HeadingFontName);
        }

        [Fact]
        public void DesignerDesignBrief_ManualFontOverridesUpdateTypographyMetadata() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-font-sync", "technical rollout proposal")
                .WithFonts("Poppins", "Aptos");

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(1);
            PowerPointDeckDesignSummary summary = brief.DescribeAlternatives(1)[0];

            Assert.Equal("Poppins", alternatives[0].Theme.HeadingFontName);
            Assert.Equal("Aptos", alternatives[0].Theme.BodyFontName);
            Assert.Equal(PowerPointTypographyStyle.Auto, alternatives[0].Theme.TypographyStyle);
            Assert.Equal(PowerPointTypographyStyle.Auto, summary.TypographyStyle);
        }

        [Fact]
        public void DesignerDesignBrief_CreativeDirectionPacksKeepAlternativesDistinct() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-pack-distinct", "technical rollout proposal")
                .WithCreativeDirectionPack(PowerPointCreativeDirectionPack.TechnicalMap);

            IReadOnlyList<PowerPointDeckDesign> alternatives = brief.CreateAlternatives(3);

            Assert.Equal(PowerPointCreativeDirectionPack.TechnicalMap, brief.CreativeDirectionPack);
            Assert.Equal(PowerPointPaletteStyle.Complementary, brief.PaletteStyle);
            Assert.Equal(PowerPointAutoLayoutStrategy.Compact, alternatives[0].BaseIntent.LayoutStrategy);
            Assert.Equal("Architecture Map", alternatives[0].Direction.Name);
            Assert.Equal("Delivery Signal", alternatives[1].Direction.Name);
            Assert.NotEqual(alternatives[0].Seed, alternatives[1].Seed);
            Assert.NotEqual(alternatives[0].Theme.Accent2Color, alternatives[1].Theme.Accent2Color);
            Assert.NotEqual(alternatives[1].Theme.Accent2Color, alternatives[2].Theme.Accent2Color);
        }

        [Fact]
        public void DesignerDesignBrief_CanRecommendDeckPlanAlternative() {
            string filePath = CreateTempPresentationPath();

            try {
                PowerPointDesignBrief brief = PowerPointDesignBrief
                    .FromBrand("#008C95", "recommended-plan", "technical rollout proposal")
                    .WithVariety(PowerPointDesignVariety.Exploratory);
                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddProcess("Rollout path", "Six steps should prefer compact geometric options.",
                        new[] {
                            new PowerPointProcessStep("Discover", "Inventory the current state."),
                            new PowerPointProcessStep("Design", "Define the target model."),
                            new PowerPointProcessStep("Pilot", "Validate the first rollout group."),
                            new PowerPointProcessStep("Deploy", "Move through controlled batches."),
                            new PowerPointProcessStep("Stabilize", "Watch early support signals."),
                            new PowerPointProcessStep("Optimize", "Tune the operating model.")
                        },
                        seed: "process")
                    .AddCoverage("Delivery coverage", null,
                        new[] {
                            new PowerPointCoverageLocation("Warsaw", 0.58, 0.43),
                            new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
                            new PowerPointCoverageLocation("Krakow", 0.58, 0.78)
                        },
                        seed: "coverage")
                    .AddCustom("Raw proof", composer => {
                        composer.AddTitle("Raw proof", "The recommendation still allows custom composition.");
                        composer.AddMetricStrip(new[] {
                            new PowerPointMetric("6", "steps"),
                            new PowerPointMetric("3", "regions")
                        }, composer.ContentArea().TakeTopCm(1.8));
                    }, seed: "raw-proof");

                IReadOnlyList<PowerPointDeckPlanAlternativeSummary> described =
                    brief.DescribeDeckPlanAlternatives(plan, 5);
                IReadOnlyList<PowerPointDeckPlanAlternativeSummary> recommended =
                    brief.RecommendDeckPlanAlternatives(plan, 5);
                PowerPointDeckPlanAlternativeSummary selected = brief.RecommendDeckPlanAlternative(plan, 5);

                Assert.Equal(described.Count, recommended.Count);
                Assert.Equal(recommended[0].Index, selected.Index);
                Assert.Equal(recommended.Max(alternative => alternative.ContentFitScore), selected.ContentFitScore);
                Assert.True(recommended[0].ContentFitScore >= recommended[recommended.Count - 1].ContentFitScore);
                Assert.Contains(selected.ContentFitReasons, reason => reason.Length > 0);
                Assert.Contains(selected.Slides, slide => slide.Kind == PowerPointDeckPlanSlideKind.Custom);

                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckComposer deck = presentation.UseDesigner(brief, plan, alternativeCount: 5);
                deck.AddSlides(plan);

                Assert.Equal(selected.Design.DirectionName, deck.Design.Direction.Name);
                List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckComposer_CanStartDirectlyFromPurposeText() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointDeckComposer deck = presentation.UseDesigner("#008C95", "proposal-client",
                    "technical rollout proposal", alternativeIndex: 1, name: "Proposal Client",
                    footerLeft: "PROPOSAL");
                deck.AddSectionSlide("Rollout plan", "Purpose-selected recipe", "cover");

                Assert.Equal("Runbook", deck.Design.Direction.Name);
                Assert.Equal(PowerPointDesignMood.Minimal, deck.Design.BaseIntent.Mood);
                Assert.Equal("Technical proposal", deck.Design.Options("cover").Eyebrow);
                Assert.Equal("PROPOSAL", deck.Design.Options("cover").FooterLeft);
                Assert.Equal(deck.Design.Theme.Name, presentation.ThemeName);
                Assert.Null(presentation.Slides[0].GetShape("Designer Direction 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckComposer_CanApplySemanticDeckPlan() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointDesignBrief brief = PowerPointDesignBrief
                    .FromBrand("#008C95", "plan-client", "technical rollout proposal")
                    .WithIdentity("Plan Client", footerLeft: "PLAN", footerRight: "Proposal");
                PowerPointDeckComposer deck = presentation.UseDesigner(brief, alternativeIndex: 0);

                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddSection("Rollout proposal", "Structured from a semantic plan.", "cover",
                        options => options.SectionVariant = PowerPointSectionLayoutVariant.EditorialRail)
                    .AddCaseStudy("Example client",
                        new[] {
                            new PowerPointCaseStudySection("Client", "A client needed clear rollout structure."),
                            new PowerPointCaseStudySection("Challenge", "The source story mixed operations and proof."),
                            new PowerPointCaseStudySection("Solution", "The plan chooses a case-study composition."),
                            new PowerPointCaseStudySection("Result", "The output remains editable.")
                        },
                        new[] {
                            new PowerPointMetric("150", "devices")
                        },
                        "case-study")
                    .AddProcess("Delivery path", "A readable rollout flow.",
                        new[] {
                            new PowerPointProcessStep("Discover", "Map constraints."),
                            new PowerPointProcessStep("Pilot", "Validate with a small group."),
                            new PowerPointProcessStep("Rollout", "Move in controlled waves.")
                        },
                        "process",
                        options => options.Variant = PowerPointProcessLayoutVariant.NumberedColumns)
                    .AddCardGrid("Scope", "Grouped workstreams.",
                        new[] {
                            new PowerPointCardContent("Deployments", new[] { "Intune", "Autopilot" }),
                            new PowerPointCardContent("Care", new[] { "Monitoring", "Reporting" })
                        },
                        "scope");

                IReadOnlyList<PowerPointSlide> slides = deck.AddSlides(plan);

                Assert.Equal(4, slides.Count);
                Assert.Equal(4, presentation.Slides.Count);
                Assert.NotNull(slides[2].GetShape("Step Marker 1"));
                Assert.Contains(slides[3].TextBoxes, textBox => textBox.Text == "PLAN");

                List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckComposer_PreflightsSemanticDeckPlanBeforeRendering() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckComposer deck = presentation.UseDesigner("#008C95", "invalid-plan",
                    PowerPointDesignRecipe.ConsultingPortfolio);
                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddSection("Cover")
                    .AddProcess("Too long", null,
                        Enumerable.Range(1, 9)
                            .Select(index => new PowerPointProcessStep("Step " + index, "Body")));
                int slideCountBefore = presentation.Slides.Count;

                PowerPointDeckPlanValidationException exception =
                    Assert.Throws<PowerPointDeckPlanValidationException>(() => deck.AddSlides(plan));

                Assert.Contains(exception.Diagnostics, diagnostic =>
                    diagnostic.Code == "Process.TooManySteps" &&
                    diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error);
                Assert.Contains("Process.TooManySteps", exception.Message);
                Assert.Equal(slideCountBefore, presentation.Slides.Count);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckComposer_CanBypassDeckPlanPreflightForLegacyBehavior() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckComposer deck = presentation.UseDesigner("#008C95", "legacy-invalid-plan",
                    PowerPointDesignRecipe.ConsultingPortfolio);
                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddProcess("Too long", null,
                        Enumerable.Range(1, 9)
                            .Select(index => new PowerPointProcessStep("Step " + index, "Body")));

                Assert.Throws<ArgumentOutOfRangeException>(() => deck.AddSlides(plan, validate: false));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckPlan_RejectsEmptySemanticContent() {
            Assert.Throws<ArgumentException>(() =>
                new PowerPointDeckPlan().AddProcess("Empty process", null, Array.Empty<PowerPointProcessStep>()));

            Assert.Throws<ArgumentNullException>(() =>
                new PowerPointDeckPlan().Add(null!));
        }

        [Fact]
        public void DesignerDeckPlan_CanMixSemanticSlidesWithRawComposition() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointDeckComposer deck = presentation.UseDesigner("#008C95", "mixed-plan",
                    PowerPointDesignRecipe.ConsultingPortfolio, footerLeft: "MIXED");
                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddSection("Mixed plan", "Semantic structure with a custom detail slide.", "cover")
                    .AddCustom("Custom detail", composer => {
                        composer.AddTitle("Custom detail", "A raw composition can live inside the plan.");
                        PowerPointLayoutBox[] columns = composer.ContentColumns(2);
                        composer.AddCardGrid(new[] {
                            new PowerPointCardContent("Choice", new[] { "Semantic plan" }),
                            new PowerPointCardContent("Escape hatch", new[] { "Raw composer" })
                        }, columns[0]);
                        composer.AddMetricStrip(new[] {
                            new PowerPointMetric("2", "modes")
                        }, columns[1].TakeTopCm(1.6));
                    }, "custom");

                IReadOnlyList<PowerPointSlide> slides = deck.AddSlides(plan);

                Assert.Equal(2, slides.Count);
                Assert.Contains(slides[1].TextBoxes, textBox => textBox.Text == "Custom detail");
                Assert.NotNull(slides[1].GetShape("Card 1"));
                Assert.NotNull(slides[1].GetShape("Metric Value 1"));

                List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckPlan_CanDescribeSlideSequenceBeforeRendering() {
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Cover", "Opening", "cover")
                .AddCaseStudy("Client",
                    new[] {
                        new PowerPointCaseStudySection("Challenge", "Many details."),
                        new PowerPointCaseStudySection("Result", "Clear story.")
                    },
                    new[] {
                        new PowerPointMetric("2", "proof points")
                    },
                    "case-study")
                .AddProcess("Path", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start."),
                        new PowerPointProcessStep("Two", "Finish.")
                    },
                    "path")
                .AddCustom("Appendix", composer => composer.AddTitle("Appendix"), "custom");

            IReadOnlyList<PowerPointDeckPlanSlideSummary> summaries = plan.DescribeSlides();

            Assert.Equal(4, summaries.Count);
            Assert.Equal(PowerPointDeckPlanSlideKind.Section, summaries[0].Kind);
            Assert.Equal("Cover", summaries[0].Title);
            Assert.Equal("cover", summaries[0].Seed);
            Assert.Equal(3, summaries[1].ContentItemCount);
            Assert.Equal(PowerPointDeckPlanSlideKind.Process, summaries[2].Kind);
            Assert.Equal(2, summaries[2].ContentItemCount);
            Assert.Equal(PowerPointDeckPlanSlideKind.Custom, summaries[3].Kind);
            Assert.Contains("Appendix", summaries[3].ToString());
        }

        [Fact]
        public void DesignerDeckPlan_CanValidateContentBeforeRendering() {
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddCaseStudy("Dense case",
                    Enumerable.Range(1, 5)
                        .Select(index => new PowerPointCaseStudySection("Section " + index, "Body")),
                    Enumerable.Range(1, 4).Select(index => new PowerPointMetric(index.ToString(), "metric")))
                .AddProcess("Long process", null,
                    Enumerable.Range(1, 9)
                        .Select(index => new PowerPointProcessStep("Step " + index, "Body")))
                .AddLogoWall("Large proof wall", null,
                    Enumerable.Range(1, 25).Select(index => new PowerPointLogoItem("Logo " + index)))
                .AddCoverage("Coverage", null,
                    Enumerable.Range(1, 19)
                        .Select(index => new PowerPointCoverageLocation("Location " + index, index == 19 ? 1.2 : 0.5, 0.5)))
                .AddCapability("Capabilities", null,
                    Enumerable.Range(1, 7).Select(index => new PowerPointCapabilitySection("Section " + index)));

            IReadOnlyList<PowerPointDeckPlanDiagnostic> diagnostics = plan.ValidateSlides();

            Assert.Contains(diagnostics, diagnostic =>
                diagnostic.Code == "CaseStudy.TooManySections" &&
                diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error);
            Assert.Contains(diagnostics, diagnostic =>
                diagnostic.Code == "CaseStudy.HiddenMetrics" &&
                diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Warning);
            Assert.Contains(diagnostics, diagnostic =>
                diagnostic.Code == "Process.TooManySteps" &&
                diagnostic.Index == 1);
            Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "LogoWall.TooManyItems");
            Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "Coverage.HiddenPins");
            Assert.Contains(diagnostics, diagnostic =>
                diagnostic.Code == "Coverage.LocationOutOfBounds" &&
                diagnostic.Message.Contains("Location 19", StringComparison.Ordinal));
            Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "Capability.TooManySections");
            Assert.Contains("Coverage.LocationOutOfBounds", diagnostics.First(diagnostic =>
                diagnostic.Code == "Coverage.LocationOutOfBounds").ToString());
        }

        [Fact]
        public void DesignerDeckPlan_ValidationAllowsReadablePlans() {
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Cover")
                .AddProcess("Process", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start."),
                        new PowerPointProcessStep("Two", "Finish.")
                    })
                .AddCapability("Capabilities", null,
                    new[] {
                        new PowerPointCapabilitySection("One", "Body"),
                        new PowerPointCapabilitySection("Two", "Body")
                    });

            Assert.Empty(plan.ValidateSlides());
        }

        [Fact]
        public void DesignerDeckComposer_PreviewsPlanSeedsFromCurrentSlidePosition() {
            string filePath = CreateTempPresentationPath();
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "seed-preview",
                    PowerPointDesignDirection.Executive);
                PowerPointDeckComposer deck = presentation.UseDesigner(design);
                deck.AddSectionSlide("Already rendered");

                PowerPointDeckPlan plan = new PowerPointDeckPlan()
                    .AddSection("Planned cover", seed: " ")
                    .AddProcess("Planned path", null,
                        new[] {
                            new PowerPointProcessStep("One", "Start."),
                            new PowerPointProcessStep("Two", "Finish.")
                        },
                        seed: " ");

                IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> neutralPreview = plan.DescribeSlides(design);
                IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> composerPreview = deck.DescribeSlides(plan);

                Assert.Equal("slide-1", neutralPreview[0].ResolvedSeed);
                Assert.Equal("slide-2", composerPreview[0].ResolvedSeed);
                Assert.Equal("slide-3", composerPreview[1].ResolvedSeed);
                Assert.Equal("seed-preview/slide-2", composerPreview[0].DesignSeed);

                IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> explicitOffsetPreview =
                    plan.DescribeSlides(design, slideIndexOffset: 1);
                Assert.Equal(composerPreview[0].ResolvedSeed, explicitOffsetPreview[0].ResolvedSeed);
                Assert.Equal(composerPreview[1].DesignSeed, explicitOffsetPreview[1].DesignSeed);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckPlan_RejectsNegativePreviewOffset() {
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "seed-preview",
                PowerPointDesignDirection.Executive);
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Planned cover");

            Assert.Throws<ArgumentOutOfRangeException>(() => plan.DescribeSlides(design, -1));
        }

        [Fact]
        public void DesignerDesignBrief_CanDescribeDeckPlanWithSlideOffset() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-offset-preview", "technical rollout proposal");
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Planned cover", seed: " ");

            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> summaries =
                brief.DescribeDeckPlan(plan, alternativeIndex: 0, slideIndexOffset: 3);

            Assert.Equal("slide-4", summaries[0].ResolvedSeed);
            Assert.Contains("/slide-4", summaries[0].DesignSeed);
        }

        [Fact]
        public void DesignerDesignBrief_CanCompareDeckPlanAlternativesBeforeRendering() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-plan-alternatives", "technical rollout proposal")
                .WithIdentity("Client", footerLeft: "CLIENT");
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Cover", "Opening", "cover")
                .AddProcess("Path", null,
                    Enumerable.Range(1, 6)
                        .Select(index => new PowerPointProcessStep("Step " + index, "Body")),
                    "path");

            IReadOnlyList<PowerPointDeckPlanAlternativeSummary> alternatives =
                brief.DescribeDeckPlanAlternatives(plan, 3);

            Assert.Equal(3, alternatives.Count);
            Assert.Equal(PowerPointDesignVariety.Balanced, alternatives[0].Variety);
            Assert.Equal("Architecture Map", alternatives[0].Design.DirectionName);
            Assert.Equal("Runbook", alternatives[1].Design.DirectionName);
            Assert.Equal("Delivery Signal", alternatives[2].Design.DirectionName);
            Assert.NotEqual(alternatives[0].Design.Accent2Color, alternatives[1].Design.Accent2Color);
            Assert.Equal(2, alternatives[0].Slides.Count);
            Assert.Equal("cover", alternatives[0].Slides[0].ResolvedSeed);
            Assert.NotEqual(alternatives[0].Slides[0].DesignSeed, alternatives[1].Slides[0].DesignSeed);
            Assert.Contains(alternatives[0].Diagnostics,
                diagnostic => diagnostic.Code == "Process.DenseSteps");
            Assert.True(alternatives[0].HasWarnings);
            Assert.False(alternatives[0].HasErrors);
            Assert.True(alternatives[0].MatchesContent);
            Assert.Equal(2, alternatives[0].ContentFitScore);
            Assert.Contains("Geometric visual style supports process, timeline, and coverage slides.",
                alternatives[0].ContentFitReasons);
            Assert.True(alternatives[1].ContentFitScore > alternatives[0].ContentFitScore);
            Assert.Contains("Compact density fits denser planned slides without manual placement.",
                alternatives[1].ContentFitReasons);
            Assert.Contains("Six or more process steps use a rail so the sequence stays connected.",
                alternatives[1].Slides[1].LayoutReasons);
            Assert.True(alternatives[2].ContentFitScore > alternatives[1].ContentFitScore);
            Assert.Contains("Architecture Map", alternatives[0].ToString());
            Assert.Contains("Balanced", alternatives[0].ToString());
            Assert.Contains("fit 2", alternatives[0].ToString());
        }

        [Fact]
        public void Composer_CreatesValidCustomSlideFromReusablePrimitives() {
            string filePath = CreateTempPresentationPath();

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointDesignTheme theme = PowerPointDesignTheme.FromBrand("2B8A6E", "Green Brand");
                    presentation.ApplyDesignerTheme(theme);
                    presentation.ComposeDesignerSlide(composer => {
                        composer.AddTitle("Custom story", "A raw composition can still use designer primitives.");
                        composer.AddCardGrid(new[] {
                            new PowerPointCardContent("Operations", new[] { "Monitoring", "Reporting" }),
                            new PowerPointCardContent("Delivery", new[] { "Rollout", "Care" })
                        }, new PowerPointCardGridSlideOptions {
                            Variant = PowerPointCardGridLayoutVariant.SoftTiles
                        });
                        composer.AddVisualFrame(PowerPointLayoutBox.FromCentimeters(20.5, 9.6, 5.6, 3.0));
                    }, theme, new PowerPointDesignerSlideOptions {
                        Eyebrow = "OfficeIMO.PowerPoint",
                        FooterLeft = "OFFICEIMO",
                        FooterRight = "Custom"
                    });

                    List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    Assert.Contains(presentation.Slides.SelectMany(slide => slide.TextBoxes),
                        textBox => textBox.Text == "Custom story");
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCoverageSlide_AddsEditablePinsAndRejectsInvalidCoordinates() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCoverageSlide("Service coverage",
                    "Pins use normalized positions inside the map panel.",
                    new[] {
                        new PowerPointCoverageLocation("Warszawa", 0.62, 0.44),
                        new PowerPointCoverageLocation("Gdansk", 0.54, 0.16),
                        new PowerPointCoverageLocation("Wroclaw", 0.34, 0.68)
                    },
                    options: new PowerPointCoverageSlideOptions {
                        Variant = PowerPointCoverageLayoutVariant.ListMap,
                        SupportingText = "Regional service teams"
                    });

                Assert.NotNull(slide.GetShape("Map Surface 1"));
                Assert.NotNull(slide.GetShape("Pin 1"));
                Assert.Equal(3, slide.Shapes.Count(shape => shape.Name != null && System.Text.RegularExpressions.Regex.IsMatch(shape.Name, @"^Pin \d+$")));
                Assert.NotNull(slide.GetShape("Location Number 1"));

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    presentation.AddDesignerCoverageSlide("Invalid", null,
                        new[] { new PowerPointCoverageLocation("Outside", 1.2, 0.4) }));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCapabilitySlide_RejectsTooManySectionsForReadableLayout() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                IEnumerable<PowerPointCapabilitySection> sections = Enumerable.Range(1, 7)
                    .Select(index => new PowerPointCapabilitySection("Section " + index));

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    presentation.AddDesignerCapabilitySlide("Too much", null, sections));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static string FormatValidationErrors(IEnumerable<OfficeOpenXmlValidationError> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.PartUri}\n" +
                    $"Path: {error.Path}"));
        }

        private static PowerPointCompositionVariant ResolveVariantForSeed(PowerPointPresentation presentation,
            string seed) {
            PowerPointCompositionVariant variant = PowerPointCompositionVariant.Auto;
            presentation.ComposeDesignerSlide(composer => {
                variant = composer.UsePreset(PowerPointCompositionPreset.MetricStory).Variant;
            }, options: new PowerPointDesignerSlideOptions {
                DesignIntent = new PowerPointDesignIntent { Seed = seed }
            });

            return variant;
        }

        private static string CreateTempPresentationPath() {
            string tempFilePath = Path.GetTempFileName();
            string presentationPath = Path.ChangeExtension(tempFilePath, ".pptx");
            File.Delete(tempFilePath);
            return presentationPath;
        }

        private static string GetTestAssetPath(string fileName) {
            return Path.Combine(RepositoryTestPaths.Find(), "Assets", fileName);
        }
    }
}
