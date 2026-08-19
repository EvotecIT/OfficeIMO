using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTableValueAnalysisTests {
    [Fact]
    public void Extract_ProfilesRichTableValueKindsForAllAdapters() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 760,
                PageHeight = 360,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Active", "Completion", "Due Date", "Start Time", "Quantity", "Owner" },
                new[] { "Yes", "25%", "2026-08-18", "09:30", "12.5", "Ada" },
                new[] { "No", "100%", "2026-08-19", "17:45", "3", "Grace" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 80, 100, 110, 100, 90, 100 }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(Assert.Single(PdfLogicalDocument.Load(pdf).Pages).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.Equal(
            new[] {
                PdfLogicalTableValueKind.Boolean,
                PdfLogicalTableValueKind.Percentage,
                PdfLogicalTableValueKind.DateTime,
                PdfLogicalTableValueKind.Time,
                PdfLogicalTableValueKind.Number,
                PdfLogicalTableValueKind.Text
            },
            data.ValueProfiles.Select(static profile => profile.Kind));
        Assert.All(data.ValueProfiles, profile => Assert.Equal(1D, profile.Confidence));
        Assert.All(data.ValueProfiles, profile => Assert.Equal(2, profile.NonEmptyCellCount));
    }

    [Fact]
    public void Analyze_MixedTypedAndTextValuesReportEvidenceBasedConfidence() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "1" },
            new[] { "2" },
            new[] { "N/A" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "Quantity" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.Text, profile.Kind);
        Assert.Equal(3, profile.NonEmptyCellCount);
        Assert.Equal(1, profile.MatchingCellCount);
        Assert.Equal(1D / 3D, profile.Confidence, 8);
    }
}
