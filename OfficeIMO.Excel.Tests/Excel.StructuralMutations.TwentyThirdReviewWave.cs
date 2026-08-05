using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_CellShifts_ClearSpansOnlyOnChangedRows() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(1, 3, "C");
            sheet.CellValue(2, 1, "Keep");
            Row changed = sheet.WorksheetPart.Worksheet.Descendants<Row>().Single(row => row.RowIndex?.Value == 1U);
            Row unchanged = sheet.WorksheetPart.Worksheet.Descendants<Row>().Single(row => row.RowIndex?.Value == 2U);
            changed.Spans = new ListValue<StringValue> { InnerText = "1:3" };
            unchanged.Spans = new ListValue<StringValue> { InnerText = "1:1" };

            sheet.InsertCells("B1:C1", ExcelCellShiftDirection.Right);

            Assert.Null(changed.Spans);
            Assert.Equal("1:1", unchanged.Spans!.InnerText);
            Assert.Equal("B", sheet.CellAt(1, 4).GetValue<string>());
            Assert.Equal("C", sheet.CellAt(1, 5).GetValue<string>());

            changed.Spans = new ListValue<StringValue> { InnerText = "1:5" };
            sheet.DeleteCells("B1:C1", ExcelCellShiftDirection.Left);

            Assert.Null(changed.Spans);
            Assert.Equal("B", sheet.CellAt(1, 2).GetValue<string>());
            Assert.Equal("C", sheet.CellAt(1, 3).GetValue<string>());
            Assert.Equal("1:1", unchanged.Spans!.InnerText);
        }

        [Fact]
        public void Test_StructuralReferenceRewrites_InvalidatePivotCacheRecordsAcrossMutationKinds() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, 10d);
            sheet.CellValue(3, 1, "West");
            sheet.CellValue(3, 2, 20d);
            sheet.AddPivotTable(
                "A1:B3",
                "F5",
                "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
            PivotTableCacheDefinitionPart cachePart = sheet.WorksheetPart.PivotTableParts.Single()
                .PivotTableCacheDefinitionPart!;

            SeedPivotCacheRecords(cachePart);
            sheet.InsertColumns(1);
            Assert.Equal("B1:C3", cachePart.PivotCacheDefinition!.CacheSource!.WorksheetSource!.Reference!.Value);
            AssertPivotCacheInvalidated(cachePart);

            SeedPivotCacheRecords(cachePart);
            sheet.InsertCells("A1:A3", ExcelCellShiftDirection.Right);
            Assert.Equal("C1:D3", cachePart.PivotCacheDefinition.CacheSource.WorksheetSource!.Reference!.Value);
            AssertPivotCacheInvalidated(cachePart);

            SeedPivotCacheRecords(cachePart);
            sheet.MoveRange("C1:D3", "A1");
            Assert.Equal("A1:B3", cachePart.PivotCacheDefinition.CacheSource.WorksheetSource!.Reference!.Value);
            AssertPivotCacheInvalidated(cachePart);
        }

        [Fact]
        public void Test_RangeMovePlan_RejectsLegacyCommentVmlAnchorOverflow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Keep");
            sheet.SetComment("A1", "Anchored note");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanMoveRange("A1", "XFD1"));

            Assert.Contains("VML anchor", exception.Message);
            Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
            Assert.Equal("Anchored note", sheet.GetComments().Single().Text);
        }

        [Fact]
        public void Test_InCellImageRemoval_ReclaimsExclusiveRichValueChain() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            byte[] image = { 1, 2, 3, 4 };

            for (int cycle = 0; cycle < 2; cycle++) {
                sheet.SetInCellImage(1, 1, image);
                Assert.True(sheet.RemoveInCellImage(1, 1));

                WorkbookPart workbookPart = document.WorkbookPartRoot;
                Assert.Empty(workbookPart.RdRichValueParts);
                Assert.Null(workbookPart.CellMetadataPart!.Metadata!.GetFirstChild<ValueMetadata>());
                Assert.DoesNotContain(workbookPart.Parts.Select(pair => pair.OpenXmlPart).OfType<ExtendedPart>(),
                    part => part.RelationshipType.Contains("richValueRel", StringComparison.Ordinal));
            }
            Assert.Empty(sheet.GetInCellImages());
            Assert.Empty(document.ValidateOpenXml());
        }

        private static void SeedPivotCacheRecords(PivotTableCacheDefinitionPart cachePart) {
            PivotCacheDefinition definition = cachePart.PivotCacheDefinition!;
            definition.RefreshOnLoad = false;
            definition.SaveData = true;
            definition.RecordCount = 1U;
            PivotTableCacheRecordsPart recordsPart = cachePart.PivotTableCacheRecordsPart
                ?? cachePart.AddNewPart<PivotTableCacheRecordsPart>();
            recordsPart.PivotCacheRecords = new PivotCacheRecords(new PivotCacheRecord()) { Count = 1U };
            recordsPart.PivotCacheRecords.Save();
            definition.Save();
        }

        private static void AssertPivotCacheInvalidated(PivotTableCacheDefinitionPart cachePart) {
            Assert.True(cachePart.PivotCacheDefinition!.RefreshOnLoad!.Value);
            Assert.False(cachePart.PivotCacheDefinition.SaveData!.Value);
            Assert.Equal(0U, cachePart.PivotCacheDefinition.RecordCount!.Value);
            Assert.Equal(0U, cachePart.PivotTableCacheRecordsPart!.PivotCacheRecords!.Count!.Value);
            Assert.Empty(cachePart.PivotTableCacheRecordsPart.PivotCacheRecords.Elements<PivotCacheRecord>());
        }
    }
}
