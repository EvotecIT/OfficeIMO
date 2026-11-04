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

        var bounded = parsed.Document.ToPowerPoint(new MarkupPowerPointExportOptions { RenderMermaidDiagrams = false }, out OfficeIMO.OfficeConversionReport conversionReport1);
        var boundedReport = (OfficeMarkupPowerPointConversionReport)conversionReport1;
        var explicitCollisionScan = parsed.Document.ToPowerPoint(new MarkupPowerPointExportOptions {
                RenderMermaidDiagrams = false,
                PreflightOptions = new PowerPointDeckPreflightOptions {
                    DetectShapeCollisions = true,
                    IncludeVisualSnapshotDiagnostics = false
                }
            }, out OfficeIMO.OfficeConversionReport conversionReport2);
        var explicitCollisionScanReport = (OfficeMarkupPowerPointConversionReport)conversionReport2;

        using (bounded)
        using (explicitCollisionScan) {
            Assert.DoesNotContain(boundedReport.Preflight.Findings,
                finding => finding.Code == "layout-shape-collision");
            Assert.Contains(explicitCollisionScanReport.Preflight.Findings,
                finding => finding.Code == "layout-shape-collision");
            Assert.Contains(explicitCollisionScanReport.Diagnostics,
                diagnostic => diagnostic.Code == "layout-shape-collision" &&
                              diagnostic.LossKind == OfficeConversionLossKind.Approximation);
            Assert.True(explicitCollisionScanReport.HasLoss);
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
        var result = parsed.Document.ToPowerPoint(new MarkupPowerPointExportOptions {
                RenderMermaidDiagrams = false,
                PreflightOptions = new PowerPointDeckPreflightOptions {
                    DetectShapeCollisions = false,
                    DetectMissingVisualAssets = false,
                    IncludeVisualSnapshotDiagnostics = false
                }
            }, out OfficeIMO.OfficeConversionReport conversionReport3);
        var resultReport = (OfficeMarkupPowerPointConversionReport)conversionReport3;

        using (result) {
            Assert.True(resultReport.Preflight.ErrorCount > 0);
            OfficeDiagnostic diagnostic = Assert.Single(
                resultReport.Diagnostics,
                item => item.Format == "OfficeIMO.Markup.PowerPoint.Preflight" &&
                        item.LossKind == OfficeConversionLossKind.Failure);
            Assert.StartsWith("slide:1", diagnostic.Location, StringComparison.Ordinal);
            Assert.True(resultReport.HasLoss);
            Assert.ThrowsAny<PowerPointDeckPreflightException>(() => resultReport.ThrowIfLoss());
        }
        using var output = new MemoryStream();
        OfficeConversionException error = Assert.Throws<OfficeConversionException>(() =>
            parsed.Document.SaveAsPowerPoint(output, new MarkupPowerPointExportOptions {
                RenderMermaidDiagrams = false,
                FailOnLoss = true,
                PreflightOptions = new PowerPointDeckPreflightOptions {
                    DetectShapeCollisions = false,
                    DetectMissingVisualAssets = false,
                    IncludeVisualSnapshotDiagnostics = false
                }
            }));
        Assert.True(Assert.IsType<OfficeMarkupPowerPointConversionReport>(error.Report).Preflight.ErrorCount > 0);
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
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
            OfficeConversionReport conversion = parsed.Document.SaveAsPowerPoint(
                presentationPath,
                new MarkupPowerPointExportOptions {
                    RenderMermaidDiagrams = false,
                    PreflightOptions = new PowerPointDeckPreflightOptions {
                        DetectShapeCollisions = false,
                        IncludeVisualSnapshotDiagnostics = false
                    }
                });
            PowerPointDeckPreflightReport report = ((OfficeMarkupPowerPointConversionReport)conversion).Preflight;
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
                    new MarkupPowerPointExportOptions {
                        RenderMermaidDiagrams = false,
                        SlideHeightInches = 1D,
                        FailOnPreflightFindings = true,
                        PreflightOptions = new PowerPointDeckPreflightOptions {
                            DetectShapeCollisions = false,
                            DetectMissingVisualAssets = false,
                            IncludeVisualSnapshotDiagnostics = false,
                            MinimumReadableFontSize = global::OfficeIMO.OfficeLength.Points(100D),
                            FailureSeverity = PowerPointDeckPreflightSeverity.Warning
                        }
                    }));

            Assert.Equal(existingOutput, File.ReadAllBytes(presentationPath));
        } finally {
            if (File.Exists(presentationPath)) File.Delete(presentationPath);
        }
    }
}
