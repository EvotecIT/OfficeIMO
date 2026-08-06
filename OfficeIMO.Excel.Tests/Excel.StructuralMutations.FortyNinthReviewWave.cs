using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImageLookup_SelectsRichValueRecordInsideMetadataBlock() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Later rich-value record");
            Metadata metadata = document.WorkbookPartRoot.CellMetadataPart!.Metadata!;
            MetadataTypes types = metadata.MetadataTypes!;
            types.Append(new MetadataType {
                Name = "OTHER",
                MinSupportedVersion = 120000U,
                Copy = true
            });
            types.Count = (uint)types.Elements<MetadataType>().Count();
            MetadataBlock block = Assert.Single(metadata.GetFirstChild<ValueMetadata>()!
                .Elements<MetadataBlock>());
            block.PrependChild(new MetadataRecord {
                TypeIndex = types.Count!.Value,
                Val = 0U
            });

            ExcelInCellImage image = Assert.Single(sheet.GetInCellImages());

            Assert.Equal("A1", image.CellReference);
            Assert.Equal("Later rich-value record", image.AltText);
            Assert.Equal(TinyPng, image.Bytes);
        }

        [Fact]
        public async Task Test_ModernChartProperties_WaitForWorkbookWriter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d }) }),
                1,
                1,
                ExcelModernChartType.Funnel)
                .SetTitle("Pipeline");

            IReadOnlyList<string?> values = await AssertWorkbookReadWaitsForWriter(
                document,
                () => new string?[] { chart.Name, chart.ChartType.ToString(), chart.Title },
                "modern chart property");

            Assert.Equal(new[] { chart.Name, chart.ChartType.ToString(), "Pipeline" }, values);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_RangeCopies_AllocateUniqueFloatingImageNames(bool transpose) {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.CellValue(1, 1, "source");
            sheet.AddImage(1, 1, TinyPng, name: "Logo");

            if (transpose) sheet.TransposeRange("A1", "B2");
            else sheet.CopyRange("A1", "B2");

            ExcelImage[] images = sheet.Images.OrderBy(image => image.RowIndex).ToArray();
            Assert.Equal(2, images.Length);
            Assert.Equal(2, images.Select(image => image.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.All(images, image => Assert.NotNull(sheet.GetImage(image.Name)));
        }

        [Fact]
        public void Test_MutationRollback_RestoresFormulaLifecycleBaselines() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Calc");
            sheet.CellValue(1, 1, 1d);
            sheet.CellFormula(1, 2, "A1+1");
            Cell formula = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            formula.CellValue = new CellValue("2");
            formula.DataType = CellValues.Number;
            formula.CellFormula!.CalculateCell = false;
            document.MarkFormulaSheetRecalculated(
                sheet.WorksheetPart,
                document.CaptureFormulaInputMutationVersion());
            Assert.False(Assert.Single(sheet.GetFormulaCells()).State.HasFlag(ExcelFormulaState.Dirty));

            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.CellValue(1, 1, 5d);
                sheet.CellFormula(1, 3, "A1+2");
                throw new InvalidOperationException("Rollback probe");
            }, new ExcelMutationPlanOptions(), CancellationToken.None));

            Assert.Equal(1d, sheet.CellAt(1, 1).GetValue<double>());
            ExcelFormulaCellInfo restored = Assert.Single(sheet.GetFormulaCells());
            Assert.Equal("B1", restored.CellReference);
            Assert.True(restored.State.HasFlag(ExcelFormulaState.Evaluated));
            Assert.False(restored.State.HasFlag(ExcelFormulaState.Dirty));
        }

        [Fact]
        public void Test_TableResize_RelocatesTotalsCommentsAndHyperlinks() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, 10d);
            sheet.AddTable("A1:B4", true, "Sales", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.TotalsRowShown = true;
            table.TotalsRowCount = 1U;
            table.AutoFilter!.Reference = "A1:B3";
            sheet.CellValue(4, 1, "Total");
            sheet.CellAt(4, 2).SetFormula("SUBTOTAL(109,Sales[Amount])");
            table.Save();
            sheet.SetComment("A4", "Legacy total", "Tester");
            sheet.AddThreadedComment("B4", "Threaded total", "Tester");
            sheet.SetHyperlinkReference(4, 1, "https://example.test/total", style: false);

            sheet.ResizeTable("Sales", "A1:B6");

            Assert.Equal("A6", Assert.Single(sheet.GetComments()).CellReference);
            Assert.Equal("B6", Assert.Single(sheet.GetThreadedComments()).CellReference);
            Assert.Equal("A6", Assert.Single(sheet.GetHyperlinks()).Key);
            Assert.False(sheet.HasComment(4, 1));
            Assert.DoesNotContain(sheet.GetThreadedComments(), comment => comment.CellReference == "B4");
            Assert.Single(sheet.WorksheetPart.HyperlinkRelationships);
        }

        [Fact]
        public void Test_LineAndScatterChartAuthoring_UsesStraightSegments() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Charts");
            var data = new ExcelChartData(
                new[] { "1", "2", "3" },
                new[] { new ExcelChartSeries("Series", new[] { 1d, 3d, 2d }) });

            sheet.AddChart(data, 1, 5, type: ExcelChartType.Line);
            sheet.AddChart(data, 18, 5, type: ExcelChartType.Scatter);

            C.LineChartSeries line = Assert.Single(sheet.WorksheetPart.DrawingsPart!.ChartParts
                .SelectMany(part => part.ChartSpace.Descendants<C.LineChartSeries>()));
            C.ScatterChartSeries scatter = Assert.Single(sheet.WorksheetPart.DrawingsPart.ChartParts
                .SelectMany(part => part.ChartSpace.Descendants<C.ScatterChartSeries>()));
            Assert.False(line.GetFirstChild<C.Smooth>()!.Val!.Value);
            Assert.False(scatter.GetFirstChild<C.Smooth>()!.Val!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
