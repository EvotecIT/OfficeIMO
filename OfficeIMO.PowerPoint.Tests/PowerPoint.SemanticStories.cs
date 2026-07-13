using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSemanticStories {
        [Fact]
        public void SemanticStoryFamilies_RenderBothVariantsAsValidEditableSlides() {
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            OfficeChartData chartData = CreateChartData();
            PowerPointArchitectureContent architecture = CreateArchitecture();
            var image = new PowerPointImageAsset(imagePath, "Product dashboard showing deployment progress") {
                Caption = "Deployment dashboard",
                Provenance = "OfficeIMO test asset",
                FocalX = 0.72,
                FocalY = 0.4
            }.Annotate(new PowerPointImageAnnotation(0.72, 0.35, "Risk", "Two sites need attention."));

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(output)) {
                    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);
                    presentation.AddDesignerExecutiveSummarySlide("Executive summary", "Metric lead",
                        CreateExecutiveContent(), options: new PowerPointExecutiveSummarySlideOptions {
                            Variant = PowerPointExecutiveSummaryLayoutVariant.MetricLead
                        });
                    presentation.AddDesignerExecutiveSummarySlide("Decision summary", "Decision brief",
                        CreateExecutiveContent(), options: new PowerPointExecutiveSummarySlideOptions {
                            Variant = PowerPointExecutiveSummaryLayoutVariant.DecisionBrief
                        });
                    presentation.AddDesignerChartStorySlide("Adoption", "Editable native chart",
                        CreateChartStory(chartData), options: new PowerPointChartStorySlideOptions {
                            Variant = PowerPointChartStoryLayoutVariant.ChartHero
                        });
                    presentation.AddDesignerChartStorySlide("Adoption signals", "Narrative rail",
                        CreateChartStory(chartData), options: new PowerPointChartStorySlideOptions {
                            Variant = PowerPointChartStoryLayoutVariant.InsightRail
                        });
                    presentation.AddDesignerComparisonSlide("Options", "Parallel decision evidence",
                        CreateComparisonItems(), options: new PowerPointComparisonSlideOptions {
                            Variant = PowerPointComparisonLayoutVariant.SideBySide
                        });
                    presentation.AddDesignerComparisonSlide("Decision matrix", "Editable table",
                        CreateComparisonItems(), options: new PowerPointComparisonSlideOptions {
                            Variant = PowerPointComparisonLayoutVariant.DecisionMatrix
                        });
                    presentation.AddDesignerScreenshotStorySlide("Product proof", "Annotated hero", image,
                        options: new PowerPointScreenshotStorySlideOptions {
                            Variant = PowerPointScreenshotStoryLayoutVariant.HeroAnnotated
                        });
                    presentation.AddDesignerScreenshotStorySlide("Product proof explained", "Narrative split", image,
                        new[] { "Deployment status remains visible.", "Risk is tied to the relevant UI region." },
                        options: new PowerPointScreenshotStorySlideOptions {
                            Variant = PowerPointScreenshotStoryLayoutVariant.SplitNarrative
                        });
                    presentation.AddDesignerAppendixTableSlide("Evidence", "Full-width editable table",
                        CreateTableData(), options: new PowerPointAppendixTableSlideOptions {
                            Variant = PowerPointAppendixTableLayoutVariant.FullWidth
                        });
                    presentation.AddDesignerAppendixTableSlide("Evidence notes", "Interpretation rail",
                        CreateTableData(withNotes: true), options: new PowerPointAppendixTableSlideOptions {
                            Variant = PowerPointAppendixTableLayoutVariant.NotesRail
                        });
                    presentation.AddDesignerArchitectureSlide("Platform", "Layered system view", architecture,
                        options: new PowerPointArchitectureSlideOptions {
                            Variant = PowerPointArchitectureLayoutVariant.Layered
                        });
                    presentation.AddDesignerArchitectureSlide("Platform relationships", "Hub-and-spoke view", architecture,
                        options: new PowerPointArchitectureSlideOptions {
                            Variant = PowerPointArchitectureLayoutVariant.HubSpoke
                        });
                    presentation.AddDesignerClosingSlide("Recommendation",
                        new PowerPointClosingContent("Start with the shared semantic core."),
                        options: new PowerPointClosingSlideOptions {
                            Variant = PowerPointClosingLayoutVariant.Statement
                        });
                    presentation.AddDesignerClosingSlide("Next action",
                        new PowerPointClosingContent("Turn evidence into action.",
                            "Inspect the proof bundle and approve the next release.", "owner@example.com"),
                        options: new PowerPointClosingSlideOptions {
                            Variant = PowerPointClosingLayoutVariant.ActionPanel
                        });

                    List<ValidationErrorInfo> errors = presentation.ValidateDocument();
                    Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
                    Assert.Equal(14, presentation.Slides.Count);
                    Assert.Equal(2, presentation.Slides.SelectMany(slide => slide.Charts).Count());
                    Assert.Equal(2, presentation.Slides.SelectMany(slide => slide.Pictures).Count());
                    Assert.True(presentation.Slides.SelectMany(slide => slide.Tables).Count() >= 3);
                    Assert.Contains(presentation.Slides.SelectMany(slide => slide.Shapes),
                        shape => shape.Name.StartsWith("Architecture Node", StringComparison.Ordinal));
                    PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                        MinimumReadableFontSizePoints = 8,
                        DetectShapeCollisions = false
                    });
                    Assert.Empty(report.Findings.Where(finding =>
                        finding.Severity == PowerPointDeckPreflightSeverity.Error));
                    presentation.Save();
                }

                using (PowerPointPresentation reopened = PowerPointPresentation.Load(output)) {
                    List<ValidationErrorInfo> errors = reopened.ValidateDocument();
                    Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
                    PowerPointPicture picture = reopened.Slides.SelectMany(slide => slide.Pictures).First();
                    Assert.Equal("Product dashboard showing deployment progress", picture.AltText);
                    Assert.True(picture.CropLeftRatio > 0 || picture.CropRightRatio > 0 ||
                        picture.CropTopRatio > 0 || picture.CropBottomRatio > 0);
                    Assert.Contains(reopened.Slides.SelectMany(slide => slide.TextBoxes),
                        textBox => textBox.Text.Contains("Source: OfficeIMO test asset"));
                }
            } finally {
                if (File.Exists(output)) File.Delete(output);
            }
        }

        [Fact]
        public void SemanticPlan_PaginatesAppendixRowsAndReportsRhythm() {
            var rows = Enumerable.Range(1, 30).Select(index =>
                (IEnumerable<string>)new[] { "Item " + index, "Ready" });
            var table = new PowerPointTableData(new[] { "Item", "Status" }, rows);
            PowerPointComparisonItem[] comparisons = CreateComparisonItems();
            var plan = new PowerPointDeckPlan()
                .AddSection("Portfolio", "Decision brief")
                .AddComparison("Choice one", null, comparisons)
                .AddComparison("Choice two", null, comparisons)
                .AddComparison("Choice three", null, comparisons)
                .AddAppendixTable("Evidence", null, table);
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#008C95", "rhythm-test");

            PowerPointDeckPlan expanded = plan.WithContinuations();
            Assert.Equal(3, expanded.Slides.Count(slide =>
                slide.Kind == PowerPointDeckPlanSlideKind.AppendixTable));
            Assert.All(expanded.Slides.OfType<PowerPointAppendixTablePlanSlide>(), slide =>
                Assert.InRange(slide.Data.Rows.Count, 1, PowerPointDeckPlanLimits.MaxAppendixTableRows));
            Assert.Empty(expanded.ValidateSlides().Where(diagnostic =>
                diagnostic.Severity == PowerPointDeckPlanDiagnosticSeverity.Error));

            PowerPointDeckRhythmReport report = plan.InspectRhythm(design);
            Assert.Contains(report.Findings, finding => finding.Code == "Rhythm.RepeatedKind");
            Assert.Contains(report.Findings, finding => finding.Code == "Rhythm.RepeatedVariant");
            Assert.Contains(report.Findings, finding => finding.Code == "Rhythm.MissingClosing");
            Assert.InRange(report.Score, 0, 99);
        }

        [Fact]
        public void SemanticPlan_ReportsMissingScreenshotBeforeRendering() {
            var image = new PowerPointImageAsset(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png"),
                "Missing test screenshot");
            var plan = new PowerPointDeckPlan().AddScreenshotStory("Proof", null, image);

            PowerPointDeckPlanDiagnostic diagnostic = Assert.Single(plan.ValidateSlides());
            Assert.Equal("ScreenshotStory.MissingImage", diagnostic.Code);
            Assert.Equal(PowerPointDeckPlanDiagnosticSeverity.Error, diagnostic.Severity);
        }

        [Theory]
        [InlineData(PowerPointImagePlacement.Fit, 0.2D, 0.2D)]
        [InlineData(PowerPointImagePlacement.Fill, 0.2D, 0.2D)]
        public void ScreenshotAnnotationsFollowFinalPictureGeometryAndCrop(
            PowerPointImagePlacement placement, double annotationX, double annotationY) {
            string imagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(imagePath, PdfPngTestImages.CreateRgbPng(2, 1));
            try {
                using var stream = new MemoryStream();
                using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
                var image = new PowerPointImageAsset(imagePath, "Geometry test image") {
                    Placement = placement
                }.Annotate(new PowerPointImageAnnotation(annotationX, annotationY, "Anchor"));

                PowerPointSlide slide = presentation.AddDesignerScreenshotStorySlide(
                    "Image geometry", null, image, options: new PowerPointScreenshotStorySlideOptions {
                        Variant = PowerPointScreenshotStoryLayoutVariant.HeroAnnotated
                    });

                PowerPointPicture picture = Assert.Single(slide.Pictures);
                PowerPointAutoShape marker = Assert.Single(slide.Shapes.OfType<PowerPointAutoShape>(),
                    shape => shape.Name == "Screenshot Annotation 1");
                double expectedX = picture.LeftCm +
                    (annotationX - picture.CropLeftRatio) /
                    (1D - picture.CropLeftRatio - picture.CropRightRatio) * picture.WidthCm;
                double expectedY = picture.TopCm +
                    (annotationY - picture.CropTopRatio) /
                    (1D - picture.CropTopRatio - picture.CropBottomRatio) * picture.HeightCm;

                Assert.InRange(Math.Abs(marker.CenterXCm - expectedX), 0D, 0.01D);
                Assert.InRange(Math.Abs(marker.CenterYCm - expectedY), 0D, 0.01D);
            } finally {
                if (File.Exists(imagePath)) File.Delete(imagePath);
            }
        }

        [Fact]
        public void ArchitectureRejectsNodeIdsThatOnlyDifferByCase() {
            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                new PowerPointArchitectureContent(new[] {
                    new PowerPointArchitectureNode("API", "Public API"),
                    new PowerPointArchitectureNode("api", "Internal API")
                }));

            Assert.Equal("nodes", exception.ParamName);
        }

        [Fact]
        public void ExecutiveSummaryRejectsDecisionPointsThatWouldBeSilentlyDropped() {
            PowerPointCardContent[] points = Enumerable.Range(1, 5)
                .Select(index => new PowerPointCardContent("Decision " + index, new[] { "Evidence" }))
                .ToArray();

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                new PowerPointExecutiveSummaryContent(Array.Empty<PowerPointMetric>(), points));

            Assert.Equal("points", exception.ParamName);
            Assert.Contains("at most four", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Theory]
        [InlineData(PowerPointExecutiveSummaryLayoutVariant.DecisionBrief, 4, 3)]
        [InlineData(PowerPointExecutiveSummaryLayoutVariant.MetricLead, 5, 4)]
        public void ExecutiveSummaryRejectsMetricsBeyondTheSelectedLayoutCapacity(
            PowerPointExecutiveSummaryLayoutVariant variant, int metricCount, int expectedCapacity) {
            var content = new PowerPointExecutiveSummaryContent(
                Enumerable.Range(1, metricCount)
                    .Select(index => new PowerPointMetric(index.ToString(), "Metric " + index)),
                Array.Empty<PowerPointCardContent>());
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                presentation.AddDesignerExecutiveSummarySlide("Summary", null, content,
                    options: new PowerPointExecutiveSummarySlideOptions { Variant = variant }));

            Assert.Equal("content", exception.ParamName);
            Assert.Contains("at most " + expectedCapacity, exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        private static PowerPointExecutiveSummaryContent CreateExecutiveContent() => new(
            new[] { new PowerPointMetric("72%", "adoption"), new PowerPointMetric("3", "decisions") },
            new[] {
                new PowerPointCardContent("Decision", new[] { "Use the shared owner" }),
                new PowerPointCardContent("Evidence", new[] { "Editable native output" })
            }, "A reusable semantic layer closes the largest quality gap.");

        private static OfficeChartData CreateChartData() => new(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new OfficeChartSeries("Adoption", new[] { 28D, 42D, 61D, 72D }),
                new OfficeChartSeries("Target", new[] { 35D, 50D, 65D, 80D })
            });

        private static PowerPointChartStoryContent CreateChartStory(OfficeChartData data) =>
            new(OfficeChartKind.ColumnClustered, data,
                new[] { "Adoption increased every quarter.", "The current gap to target is eight points." }) {
                Caption = "Quarterly product adoption",
                Provenance = "Customer success dataset",
                AlternativeText = "Clustered columns comparing quarterly adoption with target",
                DataSummary = "Adoption rose from 28 to 72 percent while target rose from 35 to 80 percent."
            };

        private static PowerPointComparisonItem[] CreateComparisonItems() => new[] {
            new PowerPointComparisonItem("Shared core", "One semantic owner",
                new[] { "Reusable", "Consistent" }, new[] { "Requires deliberate API design" }),
            new PowerPointComparisonItem("Local helpers", "Per-project composition",
                new[] { "Fast locally" }, new[] { "Drifts", "Duplicates behavior" })
        };

        private static PowerPointTableData CreateTableData(bool withNotes = false) {
            var data = new PowerPointTableData(new[] { "Capability", "Status", "Owner" }, new[] {
                new[] { "Preflight", "Ready", "PowerPoint" },
                new[] { "Templates", "Ready", "PowerPoint" },
                new[] { "Chart parity", "In progress", "Drawing" }
            }) {
                Caption = "Competitive capability evidence",
                Provenance = "OfficeIMO roadmap"
            };
            if (withNotes) data.Notes = new[] { "Output remains editable.", "Rows paginate deterministically." };
            return data;
        }

        private static PowerPointArchitectureContent CreateArchitecture() {
            PowerPointArchitectureNode[] nodes = {
                new("core", "Semantic core", "Shared contracts", "Core"),
                new("ppt", "PowerPoint", "Native authoring", "Surfaces"),
                new("markup", "Markup", "Thin adapter", "Surfaces"),
                new("export", "Export", "PNG, SVG, PDF", "Outputs")
            };
            return new PowerPointArchitectureContent(nodes, new[] {
                new PowerPointArchitectureEdge("core", "ppt", "renders"),
                new PowerPointArchitectureEdge("core", "markup", "drives"),
                new PowerPointArchitectureEdge("ppt", "export", "publishes")
            });
        }
    }
}
