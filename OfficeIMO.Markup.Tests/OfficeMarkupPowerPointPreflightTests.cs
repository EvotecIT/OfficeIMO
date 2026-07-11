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
    public void ExportWithReport_UsesSharedPowerPointPreflightAndWritesJson() {
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
            PowerPointDeckPreflightReport report = new OfficeMarkupPowerPointExporter().ExportWithReport(
                parsed.Document, new OfficeMarkupPowerPointExportOptions {
                    OutputPath = presentationPath,
                    RenderMermaidDiagrams = false,
                    PreflightReportPath = reportPath,
                    PreflightOptions = new PowerPointDeckPreflightOptions {
                        DetectShapeCollisions = false,
                        IncludeVisualSnapshotDiagnostics = false
                    }
                });

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
    public void ExportWithReport_DoesNotReplaceOutputWhenPreflightRejectsDeck() {
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
                new OfficeMarkupPowerPointExporter().ExportWithReport(parsed.Document,
                    new OfficeMarkupPowerPointExportOptions {
                        OutputPath = presentationPath,
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
