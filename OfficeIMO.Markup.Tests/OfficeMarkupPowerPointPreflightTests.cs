using System;
using System.IO;
using System.Text.Json;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Markup.Tests;

public class OfficeMarkupPowerPointPreflightTests {
    [Fact]
    public void DefaultMarkupPreflightSkipsPairwiseCollisionScanUnlessExplicitlyRequested() {
        const string markup = """
---
profile: presentation
---

# Collision policy

@slide {
  layout: blank
}

::textbox x=10% y=20% w=55% h=20%
First overlapping box

::textbox x=20% y=25% w=55% h=20%
Second overlapping box
""";
        OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup);

        OfficeMarkupPowerPointConversionResult bounded = parsed.Document.ToPowerPointPresentationResult(
            new MarkupToPowerPointOptions { RenderMermaidDiagrams = false });
        OfficeMarkupPowerPointConversionResult explicitCollisionScan = parsed.Document.ToPowerPointPresentationResult(
            new MarkupToPowerPointOptions {
                RenderMermaidDiagrams = false,
                PreflightOptions = new PowerPointDeckPreflightOptions {
                    DetectShapeCollisions = true,
                    IncludeVisualSnapshotDiagnostics = false
                }
            });

        using (bounded.Value)
        using (explicitCollisionScan.Value) {
            Assert.DoesNotContain(bounded.Report.Preflight.Findings,
                finding => finding.Code == "Layout.ShapeCollision");
            Assert.Contains(explicitCollisionScan.Report.Preflight.Findings,
                finding => finding.Code == "Layout.ShapeCollision");
            Assert.Contains(explicitCollisionScan.Report.FidelityDiagnostics,
                diagnostic => diagnostic.Code == "Layout.ShapeCollision" &&
                              diagnostic.LossKind == OfficeConversionLossKind.Approximation);
            Assert.True(explicitCollisionScan.Report.HasLoss);
        }
    }

    [Fact]
    public void PreflightFailuresParticipateInTypedLossAcceptance() {
        const string markup = """
---
profile: presentation
---

# Clipped content

@slide {
  layout: blank
}

::textbox x=10% y=120% w=55% h=20%
This shape is outside the slide.
""";
        OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup);
        OfficeMarkupPowerPointConversionResult result = parsed.Document.ToPowerPointPresentationResult(
            new MarkupToPowerPointOptions {
                RenderMermaidDiagrams = false,
                PreflightOptions = new PowerPointDeckPreflightOptions {
                    DetectShapeCollisions = false,
                    DetectMissingVisualAssets = false,
                    IncludeVisualSnapshotDiagnostics = false
                }
            });

        using (result.Value) {
            Assert.True(result.Report.Preflight.ErrorCount > 0);
            OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(
                result.Report.FidelityDiagnostics,
                item => item.Source == "OfficeIMO.Markup.PowerPoint.Preflight" &&
                        item.LossKind == OfficeConversionLossKind.Failure);
            Assert.StartsWith("slide:1", diagnostic.Location, StringComparison.Ordinal);
            Assert.True(result.Report.HasLoss);
            Assert.Throws<PowerPointDeckPreflightException>(() => result.Report.RequireNoLoss());
        }
    }

    [Fact]
    public void SaveAsPowerPoint_ReturnsSharedPreflightThatCanBePersistedSeparately() {
        const string markup = """
---
profile: presentation
title: Report Demo
---

# Delivery status

@slide {
  layout: title-and-content
}

- Scope agreed
- Delivery on track
""";
        string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        string reportPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
        try {
            OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup);
            OfficeMarkupPowerPointConversionReport conversion = parsed.Document.SaveAsPowerPoint(
                presentationPath,
                new MarkupToPowerPointOptions {
                    RenderMermaidDiagrams = false,
                    PreflightOptions = new PowerPointDeckPreflightOptions {
                        DetectShapeCollisions = false,
                        IncludeVisualSnapshotDiagnostics = false
                    }
                });
            PowerPointDeckPreflightReport report = conversion.Preflight;
            report.SaveJson(reportPath);

            Assert.True(File.Exists(presentationPath));
            Assert.True(File.Exists(reportPath));
            Assert.Equal(1, report.SlideCount);
            using JsonDocument json = JsonDocument.Parse(File.ReadAllText(reportPath));
            Assert.Equal(1, json.RootElement.GetProperty("schemaVersion").GetInt32());
            Assert.Equal(report.SlideCount, json.RootElement.GetProperty("slideCount").GetInt32());
        } finally {
            if (File.Exists(presentationPath)) File.Delete(presentationPath);
            if (File.Exists(reportPath)) File.Delete(reportPath);
        }
    }

    [Fact]
    public void SaveAsPowerPoint_DoesNotReplaceOutputWhenPreflightRejectsDeck() {
        const string markup = """
---
profile: presentation
title: Rejected export
---

# Delivery status

@slide {
  layout: title-and-content
}

A deliberately long paragraph that cannot fit inside the very short slide used by this regression test.
""";
        string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        byte[] existingOutput = { 0x45, 0x58, 0x49, 0x53, 0x54, 0x49, 0x4E, 0x47 };
        try {
            File.WriteAllBytes(presentationPath, existingOutput);
            OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup);

            Assert.Throws<PowerPointDeckPreflightException>(() =>
                parsed.Document.SaveAsPowerPoint(presentationPath,
                    new MarkupToPowerPointOptions {
                        RenderMermaidDiagrams = false,
                        SlideHeightInches = 1D,
                        FailOnPreflightFindings = true,
                        PreflightOptions = new PowerPointDeckPreflightOptions {
                            DetectShapeCollisions = false,
                            DetectMissingVisualAssets = false,
                            IncludeVisualSnapshotDiagnostics = false,
                            MinimumReadableFontSizePoints = 100D,
                            FailureSeverity = PowerPointDeckPreflightSeverity.Warning
                        }
                    }));

            Assert.Equal(existingOutput, File.ReadAllBytes(presentationPath));
        } finally {
            if (File.Exists(presentationPath)) File.Delete(presentationPath);
        }
    }
}
