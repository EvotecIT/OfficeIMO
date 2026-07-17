using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptTableWriteTests {
        [Fact]
        public void NativeWriter_RoundTripsEditableTableCellsAndLinks() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                PowerPointTable sourceTable = source
                    .AddSlide(P.SlideLayoutValues.Blank)
                    .AddTable(3, 3, PowerPointTableStylePreset.Default,
                        PowerPointUnits.FromPoints(24D),
                        PowerPointUnits.FromPoints(36D),
                        PowerPointUnits.FromPoints(300D),
                        PowerPointUnits.FromPoints(144D));
                sourceTable.GetCell(0, 0).Text = "Region";
                sourceTable.GetCell(0, 1).Text = "Revenue";
                sourceTable.GetCell(0, 1).Paragraphs[0].Runs[0].SetHyperlink(
                    "https://example.com/revenue");
                sourceTable.GetCell(0, 2).Text = "Plan";
                sourceTable.GetCell(1, 0).Text = "North";
                sourceTable.GetCell(1, 1).Text = "120";
                sourceTable.GetCell(1, 1).FillColor = "DBEAFE";
                sourceTable.GetCell(1, 2).Text = "110";
                sourceTable.GetCell(2, 0).Text = "Total";
                sourceTable.GetCell(2, 1).Text = "230";
                sourceTable.GetCell(2, 2).Text = "220";
                sourceTable.FirstRow = true;
                sourceTable.LastRow = true;
                sourceTable.FirstColumn = false;
                sourceTable.LastColumn = true;
                sourceTable.BandedRows = true;
                sourceTable.BandedColumns = false;
                sourceTable.SetCellBorders(TableCellBorders.All,
                    "C1121F", widthPoints: 1.75D);

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite);
                Assert.DoesNotContain(preflight.Findings, candidate =>
                    candidate.Code == "PPT-WRITE-TABLE-CONVERTED");
                bytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.Empty(legacy.BlipStoreEntries);
            LegacyPptShape native = Assert.Single(
                Assert.Single(legacy.Slides).Shapes,
                shape => shape.Kind == LegacyPptShapeKind.Table);
            Assert.NotNull(native.Table);
            Assert.Equal(3, native.Table!.Rows);
            Assert.Equal(3, native.Table.Columns);
            Assert.Contains(native.Table.Cells,
                cell => cell.SourceShape.Text == "230");
            Assert.True(native.Table.FirstRow);
            Assert.True(native.Table.LastRow);
            Assert.False(native.Table.FirstColumn);
            Assert.True(native.Table.LastColumn);
            Assert.True(native.Table.BandedRows);
            Assert.False(native.Table.BandedColumns);
            LegacyPptTableCell nativeCell = Assert.Single(
                native.Table.Cells, cell => cell.Row == 1
                    && cell.Column == 1);
            Assert.Equal("C1121F", nativeCell.LeftBorder?.Color);
            Assert.Equal(1.75D,
                nativeCell.LeftBorder?.WidthPoints ?? 0D, 3);

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected = PowerPointPresentation.Load(
                input);
            PowerPointTable table = Assert.Single(projected.Slides[0].Tables);
            Assert.Equal("Region", table.GetCell(0, 0).Text);
            Assert.Equal("230", table.GetCell(2, 1).Text);
            Assert.Equal("DBEAFE", table.GetCell(1, 1).FillColor);
            Assert.Equal(new Uri("https://example.com/revenue"),
                table.GetCell(0, 1).Paragraphs[0].Runs[0].Hyperlink);
            Assert.True(table.FirstRow);
            Assert.True(table.LastRow);
            Assert.False(table.FirstColumn);
            Assert.True(table.LastColumn);
            Assert.True(table.BandedRows);
            Assert.False(table.BandedColumns);
            A.LeftBorderLineProperties? leftBorder = table.GetCell(1, 1)
                .Cell.TableCellProperties?.LeftBorderLineProperties;
            Assert.Equal("C1121F", table.GetCell(1, 1).BorderColor);
            Assert.Equal((int)Math.Round(1.75D * 12700D),
                leftBorder?.Width?.Value);
            Assert.Empty(projected.Slides[0].Pictures);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void NativeWriter_BlocksMergedCellsInsteadOfFlatteningThem() {
            using PowerPointPresentation source = PowerPointPresentation.Create();
            PowerPointTable table = source.AddSlide().AddTable(2, 2);
            table.MergeCells(0, 0, 0, 1);

            LegacyPptWriteFinding finding = Assert.Single(
                source.AnalyzeLegacyPptWrite().Findings,
                candidate => candidate.Code == "PPT-WRITE-TABLE");
            Assert.Equal(LegacyPptFeature.Tables, finding.Feature);
            Assert.Contains("Merged", finding.Description);
            Assert.Throws<NotSupportedException>(() => source.ToBytes(
                PowerPointFileFormat.Ppt));
        }
    }
}
