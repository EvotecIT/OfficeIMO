using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointDeckPreflightTests {
        [Fact]
        public void Preflight_ReportsMeasuredTextOverflowAndSerializesStableJson() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBoxPoints(
                "A deliberately long paragraph that cannot fit inside a tiny text frame without losing content.",
                20, 20, 80, 14);
            text.FontSize = 20;
            text.TextAutoFit = PowerPointTextAutoFit.None;

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectShapeCollisions = false,
                IncludeVisualSnapshotDiagnostics = false
            });

            PowerPointDeckPreflightFinding finding = Assert.Single(report.Findings,
                item => item.Code == "Text.Clipped");
            Assert.Equal(PowerPointDeckPreflightSeverity.Error, finding.Severity);
            Assert.Equal(0, finding.SlideIndex);
            Assert.NotNull(finding.Bounds);
            Assert.False(report.IsSuccessful);

            string json = report.ToJson(indented: false);
            Assert.Contains("\"schemaVersion\": 1", json, StringComparison.Ordinal);
            Assert.Contains("Text.Clipped", json, StringComparison.Ordinal);
        }

        [Fact]
        public void Preflight_InspectsTextNestedInsideGroups() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox text = slide.AddTextBoxPoints(
                "A deliberately long grouped paragraph that cannot fit inside its tiny text frame.",
                20, 20, 80, 14);
            text.FontSize = 20;
            text.TextAutoFit = PowerPointTextAutoFit.None;
            PowerPointAutoShape anchor = slide.AddRectanglePoints(110, 20, 20, 20, "Group anchor");
            slide.GroupShapes(new PowerPointShape[] { text, anchor }, "Preflight group");

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectShapeCollisions = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false
            });

            Assert.Contains(report.Findings, finding =>
                finding.Code == "Text.Clipped" && finding.ShapeId == text.Id);
        }

        [Fact]
        public void Preflight_ChecksUnreadableReductionWhenOverflowDetectionIsDisabled() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointTextBox text = presentation.AddSlide().AddTextBoxPoints(
                "Auto-fit text must shrink to remain inside this deliberately tiny frame.", 20, 20, 80, 14);
            text.FontSize = 20;
            text.TextAutoFit = PowerPointTextAutoFit.Normal;

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectTextOverflow = false,
                DetectUnreadableFontReduction = true,
                MinimumReadableFontSizePoints = 19.5D,
                DetectShapeCollisions = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false
            });

            Assert.Contains(report.Findings, finding => finding.Code == "Text.UnreadableFontReduction");
            Assert.DoesNotContain(report.Findings, finding => finding.Code == "Text.Clipped");
        }

        [Fact]
        public void Preflight_ChecksCollisionsBetweenGroupedPeers() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectanglePoints(20, 20, 120, 80, "Grouped first");
            PowerPointAutoShape second = slide.AddRectanglePoints(60, 40, 120, 80, "Grouped second");
            PowerPointGroupShape group = slide.GroupShapes(new PowerPointShape[] { first, second }, "Collision group");

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectTextOverflow = false,
                DetectUnreadableFontReduction = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false,
                MinimumCollisionOverlapRatio = 0.1D
            });

            Assert.Contains(report.Findings, finding =>
                finding.Code == "Layout.ShapeCollision" && finding.ShapeIndex == 0 &&
                finding.ShapeId != group.Id);
        }

        [Fact]
        public void Preflight_ReportsOffSlideShapesAndSignificantPeerCollisions() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddRectanglePoints(20, 30, 120, 80, "First panel");
            slide.AddRectanglePoints(80, 60, 120, 80, "Second panel");
            slide.AddRectanglePoints(370, 180, 60, 60, "Off-canvas panel");

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectTextOverflow = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false,
                AllowDecorativeShapeBleed = false,
                MinimumCollisionOverlapRatio = 0.1D
            });

            Assert.Contains(report.Findings, finding => finding.Code == "Layout.ShapeCollision");
            Assert.Contains(report.Findings, finding => finding.Code == "Layout.ShapeOffSlide");
        }

        [Fact]
        public void SaveWithPreflight_RejectsSelectedSeverityBeforeSaving() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.SlideSize.SetSizePoints(400, 225);
                    PowerPointSlide slide = presentation.AddSlide();
                    presentation.Save();
                    slide.AddRectanglePoints(390, 210, 30, 30, "Outside");

                    var options = new PowerPointDeckPreflightOptions {
                        DetectTextOverflow = false,
                        DetectShapeCollisions = false,
                        DetectMissingVisualAssets = false,
                        IncludeVisualSnapshotDiagnostics = false,
                        AllowDecorativeShapeBleed = false,
                        FailureSeverity = PowerPointDeckPreflightSeverity.Error
                    };

                    PowerPointDeckPreflightException exception = Assert.Throws<PowerPointDeckPreflightException>(
                        () => presentation.SaveWithPreflight(options));
                    Assert.Contains(exception.Report.Findings,
                        finding => finding.Code == "Layout.ShapeOffSlide");
                }

                using PowerPointPresentation reopened = PowerPointPresentation.Open(path, PowerPointOpenMode.ReadOnly);
                Assert.DoesNotContain(reopened.Slides[0].Shapes, shape => shape.Name == "Outside");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Preflight_AllowsReadableContainedComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddRectanglePoints(20, 20, 360, 185, "Panel");
            PowerPointTextBox text = slide.AddTextBoxPoints("Readable content", 40, 45, 250, 50);
            text.FontSize = 18;

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false
            });

            Assert.DoesNotContain(report.Findings, finding => finding.Code == "Text.Clipped");
            Assert.DoesNotContain(report.Findings, finding => finding.Code == "Layout.ShapeCollision");
            Assert.True(report.IsSuccessful);
        }

        [Fact]
        public void Preflight_AllowsBoundedDecorativeBleedButStillRejectsContentOffSlide() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddRectanglePoints(-20, 0, 220, 225, "Intentional plane");
            PowerPointTextBox outside = slide.AddTextBoxPoints("Content outside", 390, 200, 40, 30);
            outside.Name = "Outside content";

            PowerPointDeckPreflightReport report = presentation.Preflight(new PowerPointDeckPreflightOptions {
                DetectTextOverflow = false,
                DetectShapeCollisions = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false,
                MaximumDecorativeBleedPoints = 24
            });

            Assert.DoesNotContain(report.Findings, finding => finding.ShapeName == "Intentional plane");
            Assert.Contains(report.Findings, finding => finding.ShapeName == "Outside content" &&
                finding.Code == "Layout.ShapeOffSlide");
        }
    }
}
