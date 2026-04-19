using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

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

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
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
                    Enumerable.Range(1, 7).Select(index =>
                        new PowerPointCardContent("Card " + index, new[] { "One", "Two" })),
                    options: new PowerPointCardGridSlideOptions { MaxColumns = 3 });

                PowerPointLayoutBox slideBounds = new(0, 0, presentation.SlideSize.WidthEmus, presentation.SlideSize.HeightEmus);
                for (int i = 1; i <= 7; i++) {
                    PowerPointShape? card = slide.GetShape("Designer Card " + i);
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
        public void DesignerDesignBrief_CanChooseAutoLayoutStrategy() {
            PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#008C95", "client-demo")
                .WithLayoutStrategy(PowerPointAutoLayoutStrategy.Compact);
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Opening", "Compact introduction", "cover")
                .AddCaseStudy("Client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "One story block."),
                        new PowerPointCaseStudySection("Challenge", "Second story block."),
                        new PowerPointCaseStudySection("Result", "Third story block.")
                    },
                    seed: "case")
                .AddProcess("Delivery", null,
                    new[] {
                        new PowerPointProcessStep("One", "Assess."),
                        new PowerPointProcessStep("Two", "Plan."),
                        new PowerPointProcessStep("Three", "Deliver.")
                    },
                    seed: "process")
                .AddCardGrid("Areas", null,
                    new[] {
                        new PowerPointCardContent("One"),
                        new PowerPointCardContent("Two"),
                        new PowerPointCardContent("Three")
                    },
                    seed: "cards");

            PowerPointDeckDesign design = brief.CreateDesign();
            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> summaries = brief.DescribeDeckPlan(plan);

            Assert.Equal(PowerPointAutoLayoutStrategy.Compact, brief.LayoutStrategy);
            Assert.Equal(PowerPointAutoLayoutStrategy.Compact, design.BaseIntent.LayoutStrategy);
            Assert.Equal(PowerPointAutoLayoutStrategy.Compact, design.Describe().LayoutStrategy);
            Assert.Equal("EditorialRail", summaries[0].LayoutVariant);
            Assert.Equal("EditorialSplit", summaries[1].LayoutVariant);
            Assert.Equal("NumberedColumns", summaries[2].LayoutVariant);
            Assert.Equal("AccentTop", summaries[3].LayoutVariant);
        }

        [Fact]
        public void DesignerDesignBrief_VisualLayoutStrategyPrefersHeroVariants() {
            PowerPointDesignBrief brief = PowerPointDesignBrief.FromBrand("#008C95", "client-demo")
                .WithLayoutStrategy(PowerPointAutoLayoutStrategy.VisualFirst);
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Opening", "Visual introduction", "cover")
                .AddCaseStudy("Client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "One story block."),
                        new PowerPointCaseStudySection("Result", "Second story block.")
                    },
                    seed: "case")
                .AddCoverage("Coverage", null,
                    new[] {
                        new PowerPointCoverageLocation("Warsaw", 0.54, 0.42),
                        new PowerPointCoverageLocation("Krakow", 0.56, 0.72)
                    },
                    seed: "coverage");

            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> summaries = brief.DescribeDeckPlan(plan);

            Assert.Equal("Poster", summaries[0].LayoutVariant);
            Assert.Equal("VisualHero", summaries[1].LayoutVariant);
            Assert.Equal("PinBoard", summaries[2].LayoutVariant);
        }

        [Fact]
        public void DesignerDeckComposer_CanStartFromBriefWithCustomDirections() {
            string filePath = CreateTempPresentationPath();

            try {
                PowerPointDesignBrief brief = PowerPointDesignBrief
                    .FromBrand("#008C95", "custom-brief", "executive brief")
                    .WithIdentity("Client Brief", eyebrow: "Unique deck", footerLeft: "CLIENT")
                    .WithDirections(new[] {
                        new PowerPointDesignDirection("Local Proof", PowerPointDesignMood.Editorial,
                            PowerPointSlideDensity.Relaxed, PowerPointVisualStyle.Soft, "Georgia", "Aptos",
                            showDirectionMotif: false),
                        new PowerPointDesignDirection("Operational Signal", PowerPointDesignMood.Energetic,
                            PowerPointSlideDensity.Compact, PowerPointVisualStyle.Geometric, "Poppins", "Segoe UI")
                    });

                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointDeckComposer deck = presentation.UseDesigner(brief, alternativeIndex: 1);
                deck.AddSectionSlide("Custom brief", "Caller directions beat matched recipes.", "cover");

                Assert.Equal("Operational Signal", deck.Design.Direction.Name);
                Assert.Equal(PowerPointDesignMood.Energetic, deck.Design.BaseIntent.Mood);
                Assert.Equal("Unique deck", deck.Design.Options("cover").Eyebrow);
                Assert.Equal("CLIENT", deck.Design.Options("cover").FooterLeft);
                Assert.Equal(deck.Design.Theme.Name, presentation.ThemeName);
                Assert.NotNull(presentation.Slides[0].GetShape("Designer Direction 1"));

                List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckDesign_MinimalProfileSuppressesDirectionMotifs() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "minimal-client",
                    PowerPointDesignMood.Minimal, footerLeft: "Minimal Co");

                PowerPointSlide slide = presentation.AddDesignerSectionSlide("Quiet Story", "No marker row",
                    theme: design.Theme,
                    options: design.Options("cover"));

                Assert.Equal(PowerPointVisualStyle.Minimal, design.BaseIntent.VisualStyle);
                Assert.Null(slide.GetShape("Designer Direction 1"));
                Assert.Contains(slide.TextBoxes, textBox => textBox.Text == "Minimal Co");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerDeckComposer_AppliesDesignAndKeepsRecipeOverridesSimple() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "composer-client",
                    PowerPointDesignMood.Energetic, footerLeft: "Composer Co", footerRight: "Proposal");

                PowerPointDeckComposer deck = presentation.UseDesigner(design);
                deck.AddSectionSlide("Case Study", "Project portfolio", "cover",
                    options => options.SectionVariant = PowerPointSectionLayoutVariant.EditorialRail);
                deck.AddProcessSlide("How we work", null,
                    new[] {
                        new PowerPointProcessStep("Discover", "Understand the environment."),
                        new PowerPointProcessStep("Deliver", "Ship controlled changes.")
                    },
                    "process",
                    options => options.Variant = PowerPointProcessLayoutVariant.NumberedColumns);
                deck.ComposeSlide(composer => composer.AddTitle("Custom", "Still has shared chrome."),
                    "custom", options => options.FooterRight = "Composable");

                Assert.Equal(design.Theme.Name, presentation.ThemeName);
                Assert.Equal(3, presentation.Slides.Count);
                Assert.NotNull(presentation.Slides[0].GetShape("Section Editorial Rail"));
                Assert.NotNull(presentation.Slides[1].GetShape("Process Column 1"));
                Assert.Contains(presentation.Slides[2].TextBoxes, textBox => textBox.Text == "Composable");
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
                Assert.NotNull(slides[0].GetShape("Section Editorial Rail"));
                Assert.NotNull(slides[2].GetShape("Process Column 1"));
                Assert.Contains(slides[3].TextBoxes, textBox => textBox.Text == "PLAN");

                List<ValidationErrorInfo> errors = presentation.ValidateDocument();
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
                Assert.NotNull(slides[1].GetShape("Designer Card 1"));
                Assert.NotNull(slides[1].GetShape("Composer Metric Band"));

                List<ValidationErrorInfo> errors = presentation.ValidateDocument();
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
        public void DesignerDeckPlan_CanPreviewResolvedLayoutsBeforeRendering() {
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "render-preview",
                PowerPointDesignDirection.Executive, footerLeft: "CLIENT");

            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Cover", "Opening", "cover",
                    options => options.SectionVariant = PowerPointSectionLayoutVariant.Poster)
                .AddCaseStudy("Client",
                    new[] {
                        new PowerPointCaseStudySection("Challenge", "Many details."),
                        new PowerPointCaseStudySection("Result", "Clear story.")
                    },
                    new[] {
                        new PowerPointMetric("2", "proof points")
                    },
                    "case-study",
                    options => options.Variant = PowerPointCaseStudyLayoutVariant.VisualHero)
                .AddProcess("Path", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start."),
                        new PowerPointProcessStep("Two", "Finish.")
                    },
                    "path",
                    options => options.Variant = PowerPointProcessLayoutVariant.NumberedColumns)
                .AddCardGrid("Cards", null,
                    new[] {
                        new PowerPointCardContent("One", new[] { "A" }),
                        new PowerPointCardContent("Two", new[] { "B" })
                    },
                    "cards",
                    options => options.Variant = PowerPointCardGridLayoutVariant.SoftTiles)
                .AddLogoWall("Proof", null,
                    new[] {
                        new PowerPointLogoItem("A"),
                        new PowerPointLogoItem("B")
                    },
                    "proof",
                    options => options.Variant = PowerPointLogoWallLayoutVariant.CertificateFeature)
                .AddCoverage("Coverage", null,
                    new[] {
                        new PowerPointCoverageLocation("One", 0.2, 0.3),
                        new PowerPointCoverageLocation("Two", 0.7, 0.5)
                    },
                    "coverage",
                    options => options.Variant = PowerPointCoverageLayoutVariant.ListMap)
                .AddCapability("Capabilities", null,
                    new[] {
                        new PowerPointCapabilitySection("One", "Body"),
                        new PowerPointCapabilitySection("Two", "Body")
                    },
                    "capabilities",
                    options => options.Variant = PowerPointCapabilityLayoutVariant.Stacked)
                .AddCustom("Appendix", composer => composer.AddTitle("Appendix"), "appendix", dark: true);

            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> summaries = plan.DescribeSlides(design);

            Assert.Equal(8, summaries.Count);
            Assert.Equal("Executive", summaries[0].DirectionName);
            Assert.Equal(PowerPointDesignMood.Corporate, summaries[0].Mood);
            Assert.Equal(PowerPointSlideDensity.Balanced, summaries[0].Density);
            Assert.Equal(PowerPointVisualStyle.Soft, summaries[0].VisualStyle);
            Assert.Equal("Segoe UI Semibold", summaries[0].HeadingFontName);
            Assert.Equal("Segoe UI", summaries[0].BodyFontName);
            Assert.Equal("cover", summaries[0].ResolvedSeed);
            Assert.Equal("render-preview/cover", summaries[0].DesignSeed);
            Assert.Equal("Poster", summaries[0].LayoutVariant);
            Assert.Equal("VisualHero", summaries[1].LayoutVariant);
            Assert.Equal(3, summaries[1].ContentItemCount);
            Assert.Equal("NumberedColumns", summaries[2].LayoutVariant);
            Assert.Equal("SoftTiles", summaries[3].LayoutVariant);
            Assert.Equal("CertificateFeature", summaries[4].LayoutVariant);
            Assert.Equal("ListMap", summaries[5].LayoutVariant);
            Assert.Equal("Stacked", summaries[6].LayoutVariant);
            Assert.Equal("CustomDark", summaries[7].LayoutVariant);
            Assert.Contains("Resolved process layout: NumberedColumns.", summaries[2].LayoutReasons);
            Assert.Contains("Resolved capability layout: Stacked.", summaries[6].LayoutReasons);
            Assert.Contains("The custom slide requests the dark designer surface.", summaries[7].LayoutReasons);
            Assert.Contains("Appendix", summaries[7].ToString());
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
        public void DesignerDesignBrief_CanDescribeDeckPlanBeforeRendering() {
            PowerPointDesignBrief brief = PowerPointDesignBrief
                .FromBrand("#008C95", "brief-render-preview", "technical rollout proposal")
                .WithIdentity("Client", footerLeft: "CLIENT");
            PowerPointDeckPlan plan = new PowerPointDeckPlan()
                .AddSection("Cover", "Opening", "cover")
                .AddProcess("Path", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start."),
                        new PowerPointProcessStep("Two", "Finish.")
                    },
                    "path");

            IReadOnlyList<PowerPointDeckPlanSlideRenderSummary> summaries =
                brief.DescribeDeckPlan(plan, alternativeIndex: 1);

            Assert.Equal(2, summaries.Count);
            Assert.Equal("Runbook", summaries[0].DirectionName);
            Assert.Equal(PowerPointDesignMood.Minimal, summaries[0].Mood);
            Assert.Equal("Segoe UI Semibold", summaries[0].HeadingFontName);
            Assert.Equal("cover", summaries[0].ResolvedSeed);
            Assert.Equal("brief-render-preview/technical-proposal-runbook-2/cover", summaries[0].DesignSeed);
            Assert.Equal("Rail", summaries[1].LayoutVariant);
            Assert.Equal(2, summaries[1].ContentItemCount);
            Assert.Contains("Minimal style favors a rail over heavier process decoration.",
                summaries[1].LayoutReasons);
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
        public void DesignerDeckComposer_SameContentCanUseDifferentDeckPersonalities() {
            string corporatePath = CreateTempPresentationPath();
            string minimalPath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation corporate = PowerPointPresentation.Create(corporatePath);
                corporate.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckDesign corporateDesign = PowerPointDeckDesign.FromBrand("#008C95", "same-content-a",
                    PowerPointDesignMood.Corporate);
                PowerPointDeckComposer corporateDeck = corporate.UseDesigner(corporateDesign);
                corporateDeck.AddSectionSlide("Same Story", "Different personality", "cover");

                using PowerPointPresentation minimal = PowerPointPresentation.Create(minimalPath);
                minimal.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                PowerPointDeckDesign minimalDesign = PowerPointDeckDesign.FromBrand("#008C95", "same-content-b",
                    PowerPointDesignMood.Minimal);
                PowerPointDeckComposer minimalDeck = minimal.UseDesigner(minimalDesign);
                minimalDeck.AddSectionSlide("Same Story", "Different personality", "cover");

                Assert.NotEqual(corporateDesign.Theme.HeadingFontName, minimalDesign.Theme.HeadingFontName);
                Assert.NotNull(corporate.Slides[0].GetShape("Designer Direction 1"));
                Assert.Null(minimal.Slides[0].GetShape("Designer Direction 1"));
                Assert.Equal("008C95", corporateDesign.Theme.AccentColor);
                Assert.Equal("008C95", minimalDesign.Theme.AccentColor);
            } finally {
                if (File.Exists(corporatePath)) {
                    File.Delete(corporatePath);
                }
                if (File.Exists(minimalPath)) {
                    File.Delete(minimalPath);
                }
            }
        }

        [Fact]
        public void DesignerDirectionMotif_UsesEditableTrianglesInsteadOfTextGlyphs() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerSectionSlide("Case Study", "Project portfolio");

                for (int i = 1; i <= 11; i++) {
                    PowerPointAutoShape arrow = Assert.IsAssignableFrom<PowerPointAutoShape>(
                        slide.GetShape("Designer Direction " + i));
                    Assert.Equal(A.ShapeTypeValues.Triangle, arrow.ShapeType);
                    Assert.Equal(90, arrow.Rotation);
                }

                Assert.DoesNotContain(slide.TextBoxes,
                    textBox => textBox.Text.Contains("→", StringComparison.Ordinal) ||
                               textBox.Text.Contains("➜", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerSectionSlide_CanUseEditorialRailVariant() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerSectionSlide("Case Study", "Project portfolio",
                    options: new PowerPointDesignerSlideOptions {
                        SectionVariant = PowerPointSectionLayoutVariant.EditorialRail
                    });

                Assert.NotNull(slide.GetShape("Section Editorial Rail"));
                Assert.NotNull(slide.GetShape("Section Editorial Accent Plane"));
                Assert.Null(slide.GetShape("Section Poster Frame"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCaseStudySlide_UsesPolishedVisualPlaceholderWithoutHeavyBorder() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCaseStudySlide("Example client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs a clean visual support area.")
                    });

                PowerPointShape? frame = slide.GetShape("Case Study Visual Frame");
                Assert.NotNull(frame);
                Assert.NotEqual("111111", frame!.FillColor);
                Assert.NotEqual("111111", frame.OutlineColor);
                Assert.True(frame.OutlineWidthPoints <= 0.5);
                Assert.NotNull(slide.GetShape("Case Study Visual Surface"));
                Assert.NotNull(slide.GetShape("Case Study Visual Content Panel"));
                Assert.DoesNotContain(slide.TextBoxes,
                    textBox => textBox.Text.Equals("Visual", StringComparison.OrdinalIgnoreCase));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCaseStudySlide_VisualPlaceholderVariesWithDesignIntent() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide softSlide = presentation.AddDesignerCaseStudySlide("Soft client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs a clean visual support area.")
                    },
                    options: new PowerPointCaseStudySlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Editorial, "soft-case")
                    });

                PowerPointSlide minimalSlide = presentation.AddDesignerCaseStudySlide("Minimal client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs a clean visual support area.")
                    },
                    options: new PowerPointCaseStudySlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Minimal, "minimal-case")
                    });

                Assert.NotNull(softSlide.GetShape("Visual Collage Tile 1"));
                Assert.Null(softSlide.GetShape("Case Study Visual Content Panel"));
                Assert.NotNull(minimalSlide.GetShape("Visual Diagram Node 1"));
                Assert.Null(minimalSlide.GetShape("Case Study Visual Content Panel"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCaseStudySlide_CanUseEditorialSplitVariant() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCaseStudySlide("Example client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs clean structure."),
                        new PowerPointCaseStudySection("Solution", "Use reusable sections."),
                        new PowerPointCaseStudySection("Result", "Keep the slide editable.")
                    },
                    new[] {
                        new PowerPointMetric("150", "devices")
                    },
                    options: new PowerPointCaseStudySlideOptions {
                        Variant = PowerPointCaseStudyLayoutVariant.EditorialSplit
                    });

                Assert.NotNull(slide.GetShape("Case Study Editorial Section 1"));
                Assert.NotNull(slide.GetShape("Case Study Editorial Metric Band"));
                Assert.Null(slide.GetShape("Case Study Visual Band"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCaseStudySlide_AutoUsesEditorialSplitForContentRichStory() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCaseStudySlide("Example client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs clean structure."),
                        new PowerPointCaseStudySection("Solution", "Use reusable sections."),
                        new PowerPointCaseStudySection("Result", "Keep the slide editable.")
                    },
                    new[] {
                        new PowerPointMetric("150", "devices")
                    },
                    options: new PowerPointCaseStudySlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Corporate, "auto-rich-case")
                    });

                Assert.NotNull(slide.GetShape("Case Study Editorial Section 1"));
                Assert.NotNull(slide.GetShape("Case Study Editorial Metric Band"));
                Assert.Null(slide.GetShape("Case Study Visual Band"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCaseStudySlide_CanUseVisualHeroVariant() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCaseStudySlide("Example client",
                    new[] {
                        new PowerPointCaseStudySection("Client", "Short story."),
                        new PowerPointCaseStudySection("Challenge", "Needs strong visual hierarchy."),
                        new PowerPointCaseStudySection("Solution", "Let the visual carry the left side."),
                        new PowerPointCaseStudySection("Result", "Keep supporting text compact.")
                    },
                    new[] {
                        new PowerPointMetric("150", "devices")
                    },
                    options: new PowerPointCaseStudySlideOptions {
                        Variant = PowerPointCaseStudyLayoutVariant.VisualHero
                    });

                Assert.NotNull(slide.GetShape("Case Study Visual Hero Rule"));
                Assert.NotNull(slide.GetShape("Case Study Visual Hero Metric Band"));
                Assert.NotNull(slide.GetShape("Case Study Visual Frame"));
                Assert.Null(slide.GetShape("Case Study Editorial Section 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerProcessSlide_UsesSingleRailWithDeliberateNodes() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerProcessSlide("Process", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start here."),
                        new PowerPointProcessStep("Two", "Then continue."),
                        new PowerPointProcessStep("Three", "Finish cleanly.")
                    });

                PowerPointAutoShape rail = Assert.IsAssignableFrom<PowerPointAutoShape>(
                    slide.GetShape("Process Rail"));
                Assert.Equal(A.ShapeTypeValues.Line, rail.ShapeType);
                Assert.True(rail.OutlineWidthPoints <= 1.2);

                PowerPointAutoShape firstNode = Assert.IsAssignableFrom<PowerPointAutoShape>(
                    slide.GetShape("Process Node 1"));
                PowerPointAutoShape secondNode = Assert.IsAssignableFrom<PowerPointAutoShape>(
                    slide.GetShape("Process Node 2"));
                Assert.Equal(A.ShapeTypeValues.Ellipse, firstNode.ShapeType);
                Assert.Equal(A.ShapeTypeValues.Ellipse, secondNode.ShapeType);
                Assert.True(rail.LeftCm < firstNode.RightCm);
                Assert.True(rail.RightCm > secondNode.LeftCm);
                Assert.Null(slide.GetShape("Process Arrow 1"));
                Assert.Null(slide.GetShape("Process Marker 1"));
                Assert.Null(slide.GetShape("Process Arrow Accent 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerProcessSlide_CanUseAlternateNumberedColumnVariant() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerProcessSlide("Process", null,
                    new[] {
                        new PowerPointProcessStep("One", "Start here."),
                        new PowerPointProcessStep("Two", "Then continue."),
                        new PowerPointProcessStep("Three", "Finish cleanly.")
                    },
                    options: new PowerPointProcessSlideOptions {
                        Variant = PowerPointProcessLayoutVariant.NumberedColumns
                    });

                Assert.NotNull(slide.GetShape("Process Column 1"));
                Assert.NotNull(slide.GetShape("Process Column Rule 1"));
                Assert.Null(slide.GetShape("Process Rail"));
                Assert.Null(slide.GetShape("Process Node 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerProcessSlide_AutoUsesRailForLongFlows() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerProcessSlide("Process", null,
                    Enumerable.Range(1, 6).Select(index =>
                        new PowerPointProcessStep("Step " + index, "Keep the flow readable.")),
                    options: new PowerPointProcessSlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Corporate, "auto-long-process")
                    });

                Assert.NotNull(slide.GetShape("Process Rail"));
                Assert.NotNull(slide.GetShape("Process Node 6"));
                Assert.Null(slide.GetShape("Process Column 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCardGrid_CanUseAlternateSoftTileVariant() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCardGridSlide("Services", null,
                    new[] {
                        new PowerPointCardContent("Deployments", new[] { "Intune" }),
                        new PowerPointCardContent("Maintenance", new[] { "Monitoring" })
                    },
                    options: new PowerPointCardGridSlideOptions {
                        Variant = PowerPointCardGridLayoutVariant.SoftTiles
                    });

                PowerPointShape? accent = slide.GetShape("Designer Card Accent 1");
                Assert.NotNull(accent);
                Assert.True(accent!.WidthCm < 0.25);
                Assert.True(accent.HeightCm > 1.0);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCardGrid_AutoUsesCompactAccentBarsForLargeGrids() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCardGridSlide("Services", null,
                    Enumerable.Range(1, 5).Select(index =>
                        new PowerPointCardContent("Area " + index, new[] { "One", "Two" })),
                    options: new PowerPointCardGridSlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Corporate, "auto-large-grid")
                    });

                PowerPointShape? accent = slide.GetShape("Designer Card Accent 1");
                Assert.NotNull(accent);
                Assert.True(accent!.WidthCm > 1.0);
                Assert.True(accent.HeightCm < 0.3);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCardGrid_AutoUsesSoftTilesForEditorialMood() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCardGridSlide("Services", null,
                    new[] {
                        new PowerPointCardContent("Deployments", new[] { "Intune" }),
                        new PowerPointCardContent("Maintenance", new[] { "Monitoring" })
                    },
                    options: new PowerPointCardGridSlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Editorial, "auto-soft-grid")
                    });

                PowerPointShape? accent = slide.GetShape("Designer Card Accent 1");
                Assert.NotNull(accent);
                Assert.True(accent!.WidthCm < 0.25);
                Assert.True(accent.HeightCm > 1.0);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
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

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
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
        public void Composer_CanPlacePrimitivesInsideSemanticContentRegions() {
            string filePath = CreateTempPresentationPath();

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointLayoutBox leftColumn = default;
                    PowerPointSlide slide = presentation.ComposeDesignerSlide(composer => {
                        composer.AddTitle("Structured custom slide", "Regions reduce coordinate work.");
                        PowerPointLayoutBox[] columns = composer.ContentColumns(2, gutterCm: 0.75);
                        leftColumn = columns[0];

                        composer.AddCardGrid(new[] {
                            new PowerPointCardContent("Story", new[] { "Context", "Need" }),
                            new PowerPointCardContent("Evidence", new[] { "Metrics", "Signal" })
                        }, columns[0], new PowerPointCardGridSlideOptions {
                            MaxColumns = 1,
                            Variant = PowerPointCardGridLayoutVariant.SoftTiles
                        });

                        composer.AddProcessTimeline(new[] {
                            new PowerPointProcessStep("Find", "Identify the story."),
                            new PowerPointProcessStep("Shape", "Choose structure."),
                            new PowerPointProcessStep("Finish", "Render cleanly.")
                        }, columns[1], new PowerPointProcessSlideOptions {
                            Variant = PowerPointProcessLayoutVariant.Rail
                        });

                        PowerPointLayoutBox[] lower = composer.ContentRows(2)[1].SplitColumnsCm(2, 0.55);
                        composer.AddCalloutBand("Use regions when a full recipe is too rigid.",
                            lower[0].TakeTopCm(1.35));
                        composer.AddMetricStrip(new[] {
                            new PowerPointMetric("2", "regions")
                        }, lower[1].TakeTopCm(1.35).InsetCm(0.05));
                    });

                    PowerPointShape? firstCard = slide.GetShape("Designer Card 1");
                    Assert.NotNull(firstCard);
                    Assert.True(firstCard!.Left >= leftColumn.Left);
                    Assert.True(firstCard.Right <= leftColumn.Right);
                    Assert.NotNull(slide.GetShape("Process Rail"));
                    Assert.NotNull(slide.GetShape("Composer Callout Band"));
                    Assert.NotNull(slide.GetShape("Composer Metric Band"));

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerLogoWallSlide_CanPairLogoGridWithCertificateFeature() {
            string filePath = CreateTempPresentationPath();

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointSlide slide = presentation.AddDesignerLogoWallSlide("Proof points",
                        "Reusable logo and certification wall.",
                        new[] {
                            new PowerPointLogoItem("Lenovo", "Partner"),
                            new PowerPointLogoItem("Samsung", "Devices"),
                            new PowerPointLogoItem("Brother", "Print"),
                            new PowerPointLogoItem("Epson", "Service")
                        },
                        options: new PowerPointLogoWallSlideOptions {
                            Variant = PowerPointLogoWallLayoutVariant.CertificateFeature,
                            FeatureTitle = "Featured certification"
                        });

                    Assert.NotNull(slide.GetShape("Logo Wall Tile 1"));
                    Assert.NotNull(slide.GetShape("Logo Wall Certificate Frame"));
                    Assert.NotNull(slide.GetShape("Logo Wall Certificate Document"));
                    Assert.NotNull(slide.GetShape("Logo Wall Certificate Header"));
                    Assert.NotNull(slide.GetShape("Logo Wall Certificate Seal Center"));

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerLogoWallSlide_AutoUsesCertificateFeatureWhenProofIsProvided() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerLogoWallSlide("Competence proof",
                    "Auto should choose proof emphasis when the caller supplies a featured certificate caption.",
                    new[] {
                        new PowerPointLogoItem("Xerox"),
                        new PowerPointLogoItem("Lenovo"),
                        new PowerPointLogoItem("Brother"),
                        new PowerPointLogoItem("Samsung")
                    },
                    options: new PowerPointLogoWallSlideOptions {
                        FeatureTitle = "ISO 9001 registration",
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Minimal, "proof-auto")
                    });

                Assert.NotNull(slide.GetShape("Logo Wall Tile 1"));
                Assert.NotNull(slide.GetShape("Logo Wall Certificate Frame"));
                Assert.NotNull(slide.GetShape("Logo Wall Certificate Document"));
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

                Assert.NotNull(slide.GetShape("Coverage Map Panel"));
                Assert.NotNull(slide.GetShape("Coverage Pin 1"));
                Assert.NotNull(slide.GetShape("Coverage Route 1"));
                Assert.NotNull(slide.GetShape("Coverage Map Latitude 1"));
                Assert.NotNull(slide.GetShape("Coverage List Marker 1"));

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
        public void DesignerCoverageSlide_AutoUsesListMapForManyLocations() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCoverageSlide("Coverage",
                    "A larger location set needs a readable list instead of only a pin strip.",
                    new[] {
                        new PowerPointCoverageLocation("Warszawa", 0.61, 0.45),
                        new PowerPointCoverageLocation("Gdansk", 0.55, 0.18),
                        new PowerPointCoverageLocation("Poznan", 0.38, 0.45),
                        new PowerPointCoverageLocation("Wroclaw", 0.34, 0.68),
                        new PowerPointCoverageLocation("Katowice", 0.52, 0.72),
                        new PowerPointCoverageLocation("Krakow", 0.58, 0.78),
                        new PowerPointCoverageLocation("Lublin", 0.75, 0.62)
                    },
                    options: new PowerPointCoverageSlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Corporate, "coverage-auto")
                    });

                Assert.NotNull(slide.GetShape("Coverage Map Panel"));
                Assert.NotNull(slide.GetShape("Coverage List Marker 1"));
                Assert.Null(slide.GetShape("Coverage Location Strip"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Composer_CanPlaceLogoWallAndCoverageMapInsideRegions() {
            string filePath = CreateTempPresentationPath();

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                    PowerPointSlide slide = presentation.ComposeDesignerSlide(composer => {
                        composer.AddTitle("Capability view", "Mix proof and reach without a fixed recipe.");
                        PowerPointLayoutBox[] columns = composer.ContentColumns(2, gutterCm: 0.8);
                        composer.AddLogoWall(new[] {
                            new PowerPointLogoItem("Xerox"),
                            new PowerPointLogoItem("Lenovo"),
                            new PowerPointLogoItem("Samsung"),
                            new PowerPointLogoItem("Epson")
                        }, columns[0], new PowerPointLogoWallSlideOptions {
                            MaxColumns = 2,
                            Variant = PowerPointLogoWallLayoutVariant.LogoMosaic
                        });
                        composer.AddCoverageMap(new[] {
                            new PowerPointCoverageLocation("North", 0.45, 0.18),
                            new PowerPointCoverageLocation("Central", 0.58, 0.48),
                            new PowerPointCoverageLocation("South", 0.40, 0.72)
                        }, columns[1], new PowerPointCoverageSlideOptions {
                            MapLabel = "3 regions"
                        });
                    });

                    Assert.NotNull(slide.GetShape("Logo Wall Tile 1"));
                    Assert.NotNull(slide.GetShape("Coverage Map Panel"));
                    Assert.NotNull(slide.GetShape("Coverage Pin 1"));

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, FormatValidationErrors(errors));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Composer_LogoWallAutoUsesCertificateFeatureWhenProofIsProvided() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.ComposeDesignerSlide(composer => {
                    composer.AddTitle("Proof region");
                    composer.AddLogoWall(new[] {
                        new PowerPointLogoItem("Lenovo"),
                        new PowerPointLogoItem("Samsung"),
                        new PowerPointLogoItem("Brother")
                    }, composer.ContentArea(), new PowerPointLogoWallSlideOptions {
                        FeatureTitle = "Authorized partner"
                    });
                });

                Assert.NotNull(slide.GetShape("Logo Wall Tile 1"));
                Assert.NotNull(slide.GetShape("Logo Wall Certificate Frame"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCapabilitySlide_CanCombineSectionsWithCoverageVisual() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointCapabilitySlideOptions options = new() {
                    Variant = PowerPointCapabilityLayoutVariant.TextVisual,
                    VisualKind = PowerPointCapabilityVisualKind.CoverageMap,
                    VisualLabel = "Service teams"
                };
                options.Locations.Add(new PowerPointCoverageLocation("Warszawa", 0.60, 0.48));
                options.Locations.Add(new PowerPointCoverageLocation("Gdansk", 0.55, 0.18));

                PowerPointSlide slide = presentation.AddDesignerCapabilitySlide("Service capability",
                    "Narrative sections and visual support share one composition.",
                    new[] {
                        new PowerPointCapabilitySection("Warranty service",
                            "Nationwide repair support.", new[] { "Computers", "Printers" }),
                        new PowerPointCapabilitySection("Extended care",
                            "Service beyond standard warranty.", new[] { "SLA options", "Monitoring" })
                    },
                    options: options);

                Assert.NotNull(slide.GetShape("Capability Section 1"));
                Assert.NotNull(slide.GetShape("Capability Section Accent 1"));
                Assert.NotNull(slide.GetShape("Coverage Map Panel"));
                Assert.NotNull(slide.GetShape("Coverage Pin 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCapabilitySlide_AutoUsesStackedLayoutForSectionHeavySlides() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointSlide slide = presentation.AddDesignerCapabilitySlide("Scope",
                    "Auto should keep section-heavy content structured.",
                    new[] {
                        new PowerPointCapabilitySection("Deployments", items: new[] { "Intune", "Autopilot" }),
                        new PowerPointCapabilitySection("Maintenance", items: new[] { "Monitoring", "Optimization" }),
                        new PowerPointCapabilitySection("Consulting", items: new[] { "Roadmap", "Discovery" }),
                        new PowerPointCapabilitySection("Audits", items: new[] { "Configuration", "Security review" })
                    },
                    options: new PowerPointCapabilitySlideOptions {
                        DesignIntent = PowerPointDesignIntent.FromMood(PowerPointDesignMood.Editorial, "capability-auto")
                    });

                Assert.NotNull(slide.GetShape("Capability Section 4"));
                Assert.Null(slide.GetShape("Coverage Map Panel"));
                Assert.Null(slide.GetShape("Logo Wall Tile 1"));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DesignerCapabilitySlide_CanUseStackedSectionsWithMetrics() {
            string filePath = CreateTempPresentationPath();

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

                PowerPointCapabilitySlideOptions options = new() {
                    Variant = PowerPointCapabilityLayoutVariant.Stacked
                };
                options.Metrics.Add(new PowerPointMetric("4", "areas"));
                options.Metrics.Add(new PowerPointMetric("24/7", "support"));

                PowerPointSlide slide = presentation.AddDesignerCapabilitySlide("Operations",
                    "Stacked sections work when no single visual should dominate.",
                    new[] {
                        new PowerPointCapabilitySection("Monitoring", items: new[] { "Availability", "Incidents" }),
                        new PowerPointCapabilitySection("Reporting", items: new[] { "KPIs", "Trends" }),
                        new PowerPointCapabilitySection("Optimization", items: new[] { "Backlog", "Roadmap" })
                    },
                    options: options);

                Assert.NotNull(slide.GetShape("Capability Section 3"));
                Assert.NotNull(slide.GetShape("Capability Metric Band"));
                Assert.Null(slide.GetShape("Coverage Map Panel"));
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

        private static string FormatValidationErrors(IEnumerable<ValidationErrorInfo> errors) {
            return string.Join(Environment.NewLine + Environment.NewLine,
                errors.Select(error =>
                    $"Description: {error.Description}\n" +
                    $"Id: {error.Id}\n" +
                    $"ErrorType: {error.ErrorType}\n" +
                    $"Part: {error.Part?.Uri}\n" +
                    $"Path: {error.Path?.XPath}"));
        }

        private static string CreateTempPresentationPath() {
            string tempFilePath = Path.GetTempFileName();
            string presentationPath = Path.ChangeExtension(tempFilePath, ".pptx");
            File.Delete(tempFilePath);
            return presentationPath;
        }
    }
}
