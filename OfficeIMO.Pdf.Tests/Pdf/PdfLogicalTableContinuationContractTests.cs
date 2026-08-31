using System.Globalization;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfLogicalTableContinuationContractTests {
    [Fact]
    public void TableContinuations_ExposeTypedEvidenceConfidenceAndPageScope() {
        PdfLogicalDocument document = PdfLogicalDocument.Load(BuildMultiPageTablePdf());

        PdfLogicalTableContinuationGroup group = Assert.Single(document.GetTableContinuationGroups());

        Assert.True(group.SpansPages);
        Assert.True(group.Segments.Count > 1);
        Assert.Equal(1, group.FirstPageNumber);
        Assert.Equal(group.Segments.Count, group.LastPageNumber);
        Assert.Equal(30, group.TotalRowCount);
        Assert.InRange(group.Confidence, 0.75D, 1D);
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.AdjacentPages));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.BoundaryTables));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.PageEdges));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.MatchingColumnCount));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.MatchingDetectionKind));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.CompatibleGeometry));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.CompatibleHeaders));
        Assert.True(group.Evidence.HasFlag(PdfLogicalTableContinuationEvidence.RepeatedHeaders));
    }

    [Fact]
    public void TableContinuations_CanDisableCrossPageInference() {
        PdfLogicalDocument document = PdfLogicalDocument.Load(BuildMultiPageTablePdf());

        IReadOnlyList<PdfLogicalTableContinuationGroup> groups = document.GetTableContinuationGroups(
            new PdfLogicalTableContinuationOptions { MergePageContinuations = false });

        Assert.True(groups.Count > 1);
        Assert.All(groups, group => {
            Assert.False(group.SpansPages);
            Assert.Equal(1D, group.Confidence);
            Assert.Equal(PdfLogicalTableContinuationEvidence.None, group.Evidence);
        });
    }

    [Fact]
    public void TableContinuations_PublicReaderSupportsSelectorsAndPreflight() {
        PdfDocument source = PdfDocument.Open(BuildMultiPageTablePdf());

        PdfLogicalTableContinuationGroup group = Assert.Single(source.Read.TableContinuations(PdfPageSelector.Parse("all")));
        PdfOperationResult<IReadOnlyList<PdfLogicalTableContinuationGroup>> attempt = source.Read.TryTableContinuations();

        Assert.True(group.SpansPages);
        Assert.True(attempt.Succeeded);
        Assert.Equal(PdfPreflightCapability.ReadLogicalObjects, attempt.Capability);
        Assert.True(Assert.Single(attempt.RequireValue()).SpansPages);
    }

    [Fact]
    public void TableContinuations_RejectInvalidConfidence() {
        PdfLogicalDocument document = PdfLogicalDocument.Load(BuildMultiPageTablePdf());

        Assert.Throws<ArgumentOutOfRangeException>(() => document.GetTableContinuationGroups(
            new PdfLogicalTableContinuationOptions { MinimumConfidence = double.NaN }));
    }

    private static byte[] BuildMultiPageTablePdf() {
        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int index = 1; index <= 30; index++) {
            rows.Add(new[] {
                "Check " + index.ToString(CultureInfo.InvariantCulture),
                "Team " + index.ToString(CultureInfo.InvariantCulture)
            });
        }

        return PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(rows, style: new PdfTableStyle {
                HeaderRowCount = 2,
                RepeatHeaderRowCount = 2,
                ColumnWidthPoints = new List<double?> { 120, 120 },
                CellPaddingX = 5,
                CellPaddingY = 3
            })
            .ToBytes();
    }
}
