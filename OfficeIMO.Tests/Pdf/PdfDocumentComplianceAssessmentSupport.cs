using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    private static OfficeShape CreateComplianceShape(double width = 24, double height = 24) {
        OfficeShape shape = OfficeShape.Rectangle(width, height);
        shape.FillColor = OfficeColor.FromRgb(219, 234, 254);
        shape.StrokeColor = OfficeColor.FromRgb(37, 99, 235);
        shape.StrokeWidth = 1;
        return shape;
    }

    private static PdfDocument CreatePdfA3GroundworkDocument() {
        return PdfDocument.Create(CreatePdfA3GroundworkOptions());
    }

    private static PdfOptions CreatePdfA3GroundworkOptions() {
        return new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();
    }

    private static PdfOptions CreatePdfUaGroundworkOptions() {
        return new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true,
                Language = "en-US",
                ViewerPreferences = new PdfViewerPreferencesOptions {
                    DisplayDocTitle = true
                }
            }
            .SetPdfUaIdentification();
    }

    private static PdfComplianceRequirement AssertRequirement(PdfComplianceReadinessReport report, string id, PdfComplianceRequirementStatus status) {
        PdfComplianceRequirement requirement = Assert.Single(report.Requirements, requirement => requirement.Id == id);
        Assert.Equal(status, requirement.Status);
        Assert.False(string.IsNullOrWhiteSpace(requirement.DisplayName));
        Assert.False(string.IsNullOrWhiteSpace(requirement.Diagnostic));
        return requirement;
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

}
