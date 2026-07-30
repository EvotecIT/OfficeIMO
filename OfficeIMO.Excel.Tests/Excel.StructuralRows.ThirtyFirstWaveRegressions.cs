using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesUnchangedFormulaRecalculation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellFormula(1, 1, "1+1");
            CellFormula formula = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<CellFormula>());
            formula.CalculateCell = null;

            ExcelRowMutationPlan plan = data.PlanInsertRows(100);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);

            plan.Apply();

            Assert.True(formula.CalculateCell?.Value);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesPendingUnchangedFormulaRecalculation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet summary = document.AddWorksheet("Summary");
            for (int column = 1; column <= 128; column++) {
                summary.CellValue(1, column, column);
            }
            summary.CellFormula(1, 129, "1+1");
            Assert.True(document.HasPendingDirectCellValues);
            ExcelSheet data = AddWorksheetWithoutMaterializingPending(
                document,
                "Data");

            ExcelRowMutationPlan plan = data.PlanInsertRows(100);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);

            plan.Apply();

            CellFormula formula = Assert.Single(
                summary.WorksheetPart.Worksheet.Descendants<CellFormula>());
            Assert.True(formula.CalculateCell?.Value);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresUnchangedHyperlinksAndTables() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Name");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.SetInternalLink(3, 1, sheet, "A1", display: "Top");
            sheet.AddTable(
                "A1:B2",
                hasHeader: true,
                name: "DataTable",
                OfficeIMO.Excel.TableStyle.TableStyleMedium2);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(100);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "hyperlinks");
            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "tables");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresUnchangedPivotOutputAndSource() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(100);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "pivots");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesIndependentCommentVmlAnchorMove() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment("A1", "Anchored note");
            var vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            string before;
            using (var reader = new StreamReader(vmlPart.GetStream())) {
                before = reader.ReadToEnd();
            }

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(3);

            ExcelMutationImpact comments = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "comments");
            Assert.Equal(1, comments.ItemCount);
            plan.Apply();

            string after;
            using (var reader = new StreamReader(vmlPart.GetStream())) {
                after = reader.ReadToEnd();
            }
            Assert.NotEqual(before, after);
            Assert.True(sheet.HasComment(1, 1));
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIgnoresVmlShapeWithInvalidCoordinates() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment("A1", "Malformed coordinate");
            var vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XNamespace excelNamespace =
                "urn:schemas-microsoft-com:office:excel";
            XDocument vml;
            using (Stream stream = vmlPart.GetStream()) {
                vml = XDocument.Load(stream);
            }
            vml.Descendants(excelNamespace + "Row").Single().Value = "invalid";
            using (Stream stream = vmlPart.GetStream(FileMode.Create)) {
                vml.Save(stream, SaveOptions.DisableFormatting);
            }

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(3);

            Assert.DoesNotContain(
                plan.Impacts,
                impact => impact.Category == "comments");
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsVmlShapeWithoutObjectType() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment("A1", "Missing object type");
            var vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XNamespace excelNamespace =
                "urn:schemas-microsoft-com:office:excel";
            XDocument vml;
            using (Stream stream = vmlPart.GetStream()) {
                vml = XDocument.Load(stream);
            }
            vml.Descendants(excelNamespace + "ClientData")
                .Single()
                .Attribute("ObjectType")
                ?.Remove();
            using (Stream stream = vmlPart.GetStream(FileMode.Create)) {
                vml.Save(stream, SaveOptions.DisableFormatting);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanInsertRows(3));

            Assert.Contains(
                "form controls",
                exception.Message,
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
