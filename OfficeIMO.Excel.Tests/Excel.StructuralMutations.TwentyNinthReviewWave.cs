using System;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ReferenceAlgebra_NormalizesEquivalentQuotedThreeDimensionalQualifiers() {
            ExcelReference combined = ExcelReference.Parse("'First Sheet:Last Sheet'!A1:B2");
            ExcelReference separate = ExcelReference.Parse("'First Sheet':'Last Sheet'!B2:C3");

            Assert.True(combined.Intersects(separate));
            Assert.Equal("'First Sheet:Last Sheet'!B2", combined.Intersect(separate)!.ToString());
            Assert.Equal("'First Sheet:Last Sheet'!A1:C3", combined.BoundingUnion(separate).ToString());
            Assert.Equal(
                ExcelReference.Parse("'First Sheet':'Last Sheet'!A1:B2"),
                combined);
            Assert.Equal(
                ExcelReference.Parse("'First Sheet':'Last Sheet'!A1:B2").GetHashCode(),
                combined.GetHashCode());

            Assert.Equal(
                ExcelReference.Parse("'First''s:Last''s'!A1"),
                ExcelReference.Parse("'First''s':'Last''s'!A1"));
            Assert.Equal(
                ExcelReference.Parse("'[Book.xlsx]First Sheet:[Book.xlsx]Last Sheet'!A1"),
                ExcelReference.Parse("'[Book.xlsx]First Sheet':'[Book.xlsx]Last Sheet'!A1"));
        }

        [Fact]
        public void Test_RemoveWorksheet_PreflightsRelationshipBeforePivotInteractionCleanup() {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sales = document.AddWorksheet("Sales");
            sales.CellValue(1, 1, "Region");
            sales.CellValue(1, 2, "Sales");
            sales.CellValue(2, 1, "East");
            sales.CellValue(2, 2, 10d);
            sales.AddPivotTable(
                sourceRange: "A1:B2",
                destinationCell: "D2",
                name: "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) });
            document.AddPivotSlicer(
                "SalesPivot",
                "Region",
                "Sales",
                new ExcelSlicerViewOptions { Name = "RegionFilter" });
            Assert.Single(document.GetPivotInteractions());

            string relationshipId = sales.SheetElement.Id!.Value!;
            sales.SheetElement.Id = "missingRelationship";
            try {
                Assert.Throws<InvalidDataException>(() => document.RemoveWorksheet("Sales"));
                Assert.Single(document.GetPivotInteractions());
                Assert.Contains(document.Sheets, sheet => sheet.Name == "Sales");
            } finally {
                sales.SheetElement.Id = relationshipId;
            }

            document.RemoveWorksheet("Sales");
            Assert.Empty(document.GetPivotInteractions());
            Assert.DoesNotContain(document.Sheets, sheet => sheet.Name == "Sales");
        }
    }
}
