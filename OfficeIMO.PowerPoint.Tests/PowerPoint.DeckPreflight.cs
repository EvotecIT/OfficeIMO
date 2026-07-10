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
            PowerPointSlide slide = presentation.Slides[0];
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
        public void Preflight_ReportsOffSlideShapesAndSignificantPeerCollisions() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.Slides[0];
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
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            presentation.Slides[0].AddRectanglePoints(390, 210, 30, 30, "Outside");

            var options = new PowerPointDeckPreflightOptions {
                DetectTextOverflow = false,
                DetectShapeCollisions = false,
                DetectMissingVisualAssets = false,
                IncludeVisualSnapshotDiagnostics = false,
                AllowDecorativeShapeBleed = false,
                FailureSeverity = PowerPointDeckPreflightSeverity.Error
            };

            PowerPointDeckPreflightException exception =
                Assert.Throws<PowerPointDeckPreflightException>(() => presentation.SaveWithPreflight(options));
            Assert.Contains(exception.Report.Findings, finding => finding.Code == "Layout.ShapeOffSlide");
        }

        [Fact]
        public void Preflight_AllowsReadableContainedComposition() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(400, 225);
            PowerPointSlide slide = presentation.Slides[0];
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
            PowerPointSlide slide = presentation.Slides[0];
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
