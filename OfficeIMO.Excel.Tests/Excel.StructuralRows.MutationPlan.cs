using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Xm = DocumentFormat.OpenXml.Office.Excel;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MutationPlanIsNonMutatingAndAppliesOnce() {
            using var document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(1, 1, "Header");
            data.CellValue(2, 1, 10);
            data.CellValue(3, 1, 20);
            data.CellFormula(1, 3, "SUM(A2:A3)");
            summary.CellFormula(1, 1, "'Data'!A3");
            document.SetNamedRange("DataRows", "'Data'!A2:A3", save: false);

            ExcelRowMutationPlan plan = data.PlanInsertRows(2);

            Assert.False(plan.IsConsumed);
            Assert.Equal(ExcelRowMutationKind.Insert, plan.Kind);
            Assert.Equal("Data", plan.SheetName);
            Assert.Equal(2, plan.FirstRow);
            Assert.Equal(1, plan.Count);
            Assert.True(plan.RequiresFullRecalculation);
            Assert.True(plan.ScannedElements > 0);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "worksheet-cells" && impact.ItemCount == 2);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount >= 2);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "defined-names" && impact.ItemCount == 1);
            Assert.Equal(10, data.CellAt(2, 1).GetValue<int>());

            plan.Apply();

            Assert.True(plan.IsApplied);
            Assert.True(plan.IsConsumed);
            Assert.True(data.CellAt(2, 1).GetValue().IsBlank);
            Assert.Equal(10, data.CellAt(3, 1).GetValue<int>());
            Assert.Equal("SUM(A3:A4)", data.GetFormulaText(1, 3));
            Assert.Equal("'Data'!A4", summary.GetFormulaText(1, 1));
            Assert.Throws<InvalidOperationException>(() => plan.Apply());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanHonorsInspectionBudgetWithoutMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            for (int row = 1; row <= 10; row++) {
                sheet.CellValue(row, 1, row);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanDeleteRows(
                    2,
                    options: new ExcelMutationPlanOptions { MaximumScannedElements = 3 }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
            Assert.Equal(10, sheet.CellAt(10, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsUnloadedLargeSheetBeforeMaterializingItsDom() {
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.MutationBudget.{Guid.NewGuid():N}.xlsx");
            try {
                using (var source = ExcelDocument.Create(path)) {
                    ExcelSheet target = source.AddWorksheet("Target");
                    target.CellValue(1, 1, "Value");
                    ExcelSheet unrelated = source.AddWorksheet("Unrelated");
                    for (int row = 1; row <= 100; row++) {
                        unrelated.CellValue(row, 1, row);
                    }
                    source.Save();
                }

                using var document = ExcelDocument.Load(path);
                Sheet unrelatedSheet = document.WorkbookRoot.Sheets!
                    .Elements<Sheet>()
                    .Single(sheet => sheet.Name?.Value == "Unrelated");
                WorksheetPart unrelatedPart = (WorksheetPart)document.WorkbookPartRoot
                    .GetPartById(unrelatedSheet.Id!);
                Assert.False(unrelatedPart.IsRootElementLoaded);

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                    document["Target"].PlanInsertRows(
                        1,
                        options: new ExcelMutationPlanOptions { MaximumScannedElements = 50 }));

                Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
                Assert.False(unrelatedPart.IsRootElementLoaded);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRejectsDeferredWritesWithoutMaterializing() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            var table = new DataTable("Rows");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Score", typeof(int));
            table.Rows.Add("North", 10);
            table.Rows.Add("South", 20);
            sheet.InsertDataTable(table);
            Assert.True(document.HasDeferredDirectDataSetImport);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertRows(2));

            Assert.Contains("pending deferred", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(document.HasDeferredDirectDataSetImport);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsCommentVmlBeforeMaterializingIt() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue("Commented");
            sheet.SetComment(2, 1, "Review", author: "Tester");
            int baseline = sheet.PlanInsertRows(2).ScannedElements;

            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XDocument vml;
            using (Stream source = vmlPart.GetStream()) {
                vml = XDocument.Load(source);
            }
            XNamespace vmlNamespace = "urn:schemas-microsoft-com:vml";
            for (int index = 0; index < 1_000; index++) {
                vml.Root!.Add(
                    new XElement(
                        vmlNamespace + "shape",
                        new XAttribute("id", $"_budget_shape_{index}")));
            }
            using (Stream destination = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vml.Save(destination);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertRows(
                    2,
                    options: new ExcelMutationPlanOptions {
                        MaximumScannedElements = baseline + 20
                    }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal("Commented", sheet.CellAt(2, 1).GetValue<string>());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanAppliesCharacterBudgetToUnloadedVml() {
            string path = Path.Combine(
                Path.GetTempPath(),
                $"OfficeIMO.Excel.MutationVmlCharacterBudget.{Guid.NewGuid():N}.xlsx");
            try {
                using (var source = ExcelDocument.Create(path)) {
                    ExcelSheet sourceSheet = source.AddWorksheet("Data");
                    sourceSheet.CellAt(2, 1).SetValue("Commented");
                    sourceSheet.SetComment(2, 1, "Review", author: "Tester");
                    source.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                    VmlDrawingPart vmlPart = Assert.Single(
                        package.WorkbookPart!.WorksheetParts.Single().VmlDrawingParts);
                    XDocument vml;
                    using (Stream source = vmlPart.GetStream()) {
                        vml = XDocument.Load(source);
                    }
                    vml.Root!.SetAttributeValue(
                        "data-mutation-budget-padding",
                        new string('x', 100_000));
                    using Stream destination = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
                    vml.Save(destination);
                }

                using var document = ExcelDocument.Load(path);
                ExcelSheet sheet = document["Data"];
                VmlDrawingPart unloadedVmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
                Assert.False(unloadedVmlPart.IsRootElementLoaded);

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                    sheet.PlanInsertRows(
                        2,
                        options: new ExcelMutationPlanOptions {
                            MaximumScannedElements = 10_000,
                            MaximumScannedCharacters = 64_000
                        }));

                Assert.Contains("decompressed XML bytes", exception.Message, StringComparison.Ordinal);
                Assert.False(unloadedVmlPart.IsRootElementLoaded);
                Assert.Equal("Commented", sheet.CellAt(2, 1).GetValue<string>());
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanPreservesPendingDirectCellBuffer() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            Assert.True(document.HasPendingDirectCellValues);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            Assert.True(document.HasPendingDirectCellValues);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "worksheet-cells" && impact.ItemCount == 2);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsMalformedSheetEntriesBeforeRelationshipLookup() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Sheets sheets = Assert.IsType<Sheets>(document.WorkbookRoot.Sheets);
            for (uint index = 0; index < 10; index++) {
                sheets.Append(new Sheet {
                    Name = "Malformed" + index,
                    SheetId = 100U + index
                });
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanDeleteRows(
                    1,
                    options: new ExcelMutationPlanOptions { MaximumScannedElements = 5 }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRevalidatesBeforeApply() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);
            sheet.SetArrayFormula("B1:B3", "A1:A3*2");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => plan.Apply());

            Assert.Contains("array formula", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(plan.IsApplied);
            Assert.True(plan.IsConsumed);
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
            InvalidOperationException retry = Assert.Throws<InvalidOperationException>(() => plan.Apply());
            Assert.Contains("previous attempt failed", retry.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanRespectsDefinedNameScope() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            var globalName = new DefinedName("'Data'!$A$5:$B$7") {
                Name = "DataRows"
            };
            var otherLocalName = new DefinedName("$A$2:$B$3") {
                Name = "OtherRows",
                LocalSheetId = 1U
            };
            document.WorkbookRoot.DefinedNames = new DefinedNames(
                globalName,
                otherLocalName);

            ExcelRowMutationPlan plan = data.PlanDeleteRows(2);

            ExcelMutationImpact names = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "defined-names");
            Assert.Equal(1, names.ItemCount);
            Assert.Equal("'Data'!$A$5:$B$7", globalName.Text);
            Assert.Equal("$A$2:$B$3", otherLocalName.Text);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanUsesImplicitCellCoordinates() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Header");
            sheet.CellAt(2, 1).SetValue(10);
            sheet.CellAt(3, 1).SetValue(20);
            foreach (Row row in sheet.WorksheetPart.Worksheet.Descendants<Row>()) {
                row.RowIndex = null;
                foreach (Cell cell in row.Elements<Cell>()) {
                    cell.CellReference = null;
                }
            }

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact cells = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "worksheet-cells");
            Assert.Equal(2, cells.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanResolvesSharedFormulaFollowers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            sheet.CellAt(3, 1).SetValue(2);
            AppendSharedFormulaGroup(sheet, sharedIndex: 41U, anchorReference: "B2:B3");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(2, formulas.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanUsesPendingFormulaAsAuthoritative() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "A2");
            Assert.True(document.HasPendingDirectCellValues);

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesThreadedComments() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(3, 1).SetValue("Value");
            sheet.AddThreadedComment("A3", "Review");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact comments = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "comments");
            Assert.Equal(1, comments.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesRepliesRemovedWithDeletedParent() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            ExcelThreadedCommentResult parent =
                sheet.AddThreadedComment("A3", "Parent");
            sheet.ReplyToThreadedComment(parent.Id, "Reply");

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(3);

            ExcelMutationImpact comments = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "comments");
            Assert.Equal(2, comments.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanTraversesReverseOrderedThreadedCommentChain() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            WorksheetThreadedCommentsPart threadedPart =
                sheet.WorksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
            var threadedComments = new Threaded.ThreadedComments();
            const int commentCount = 128;
            string[] ids = Enumerable.Range(0, commentCount)
                .Select(index => $"{{00000000-0000-0000-0000-{index:D12}}}")
                .ToArray();
            for (int index = commentCount - 1; index >= 0; index--) {
                var comment = new Threaded.ThreadedComment(
                    new Threaded.ThreadedCommentText($"Comment {index}")) {
                    Id = ids[index],
                    Ref = index == 0 ? "A3" : null,
                    ParentId = index == 0 ? null : ids[index - 1]
                };
                threadedComments.Append(comment);
            }
            threadedPart.ThreadedComments = threadedComments;

            ExcelRowMutationPlan plan = sheet.PlanDeleteRows(3);

            ExcelMutationImpact comments = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "comments");
            Assert.Equal(commentCount, comments.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsThreadedCommentsBeforeCollectingThem() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            WorksheetThreadedCommentsPart threadedPart =
                sheet.WorksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
            var threadedComments = new Threaded.ThreadedComments();
            for (int index = 0; index < 128; index++) {
                threadedComments.Append(new Threaded.ThreadedComment(
                    new Threaded.ThreadedCommentText($"Comment {index}")) {
                    Id = $"{{00000000-0000-0000-0000-{index:D12}}}",
                    Ref = "A3"
                });
            }
            threadedPart.ThreadedComments = threadedComments;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanDeleteRows(
                    3,
                    options: new ExcelMutationPlanOptions {
                        MaximumScannedElements = 64
                    }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(128, threadedPart.ThreadedComments.Elements<Threaded.ThreadedComment>().Count());
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesChartsHostedOnOtherSheets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue(10);
            data.CellAt(6, 1).SetValue(20);
            summary.CellAt(1, 1).SetValue("Label");
            summary.CellAt(1, 2).SetValue("Value");
            summary.CellAt(2, 1).SetValue("One");
            summary.CellAt(2, 2).SetValue(1);
            summary.AddChartFromRange("A1:B2", row: 5, column: 4);
            ChartPart chartPart = Assert.Single(summary.WorksheetPart.DrawingsPart!.ChartParts);
            C.Formula formula = chartPart.ChartSpace.Descendants<C.Formula>().First();
            formula.Text = "Data!A5:A6";

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount >= 1);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "drawings" && impact.ItemCount >= 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesPivotSourcesHostedOnOtherSheets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = CreatePivotSheet(document);
            ExcelSheet summary = document.AddWorksheet("Summary");
            PivotTablePart pivotPart = Assert.Single(data.WorksheetPart.PivotTableParts);
            summary.WorksheetPart.AddPart(pivotPart);
            data.WorksheetPart.DeletePart(pivotPart);

            ExcelRowMutationPlan plan = data.PlanInsertRows(2);

            ExcelMutationImpact pivots = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "pivots");
            Assert.True(pivots.ItemCount >= 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanBudgetsEveryPivotDefinitionElement() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            PivotTablePart pivotPart = Assert.Single(sheet.WorksheetPart.PivotTableParts);
            PivotFields pivotFields = Assert.IsType<PivotFields>(
                pivotPart.PivotTableDefinition!.GetFirstChild<PivotFields>());
            for (int index = 0; index < 256; index++) {
                pivotFields.Append(new PivotField());
            }
            pivotFields.Count = (uint)pivotFields.ChildElements.Count;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertRows(
                    2,
                    options: new ExcelMutationPlanOptions {
                        MaximumScannedElements = 128
                    }));

            Assert.Contains("exceeded its limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(258, pivotFields.ChildElements.Count);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanCountsFormulaCellOnceAcrossReferenceAttributes() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 3).SetValue(0);
            Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(candidate => candidate.CellReference?.Value == "C2");
            cell.CellFormula = new CellFormula("1") {
                FormulaType = CellFormulaValues.DataTable,
                Reference = "C2:C3",
                R1 = "A2",
                R2 = "B3"
            };

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            ExcelMutationImpact formulas = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "formula-references");
            Assert.Equal(1, formulas.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesCrossSheetHyperlinks() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellAt(5, 1).SetValue("Target");
            summary.SetInternalLink(1, 1, data, "A5", display: "Target");

            ExcelRowMutationPlan plan = data.PlanInsertRows(5);

            ExcelMutationImpact hyperlinks = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "hyperlinks");
            Assert.Equal(1, hyperlinks.ItemCount);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "formula-references" && impact.ItemCount >= 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesOffice2010ValidationAndFormatting() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            var validation = new X14.DataValidation(
                new Xm.ReferenceSequence("B2:B3"));
            var formatting = new X14.ConditionalFormatting(
                new Xm.ReferenceSequence("C2:C3"));
            sheet.WorksheetPart.Worksheet.Append(
                new ExtensionList(
                    new Extension(validation) {
                        Uri = "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"
                    },
                    new Extension(formatting) {
                        Uri = "{78C0D931-6437-407D-A8EE-F0AAD7539E65}"
                    }));

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(2);

            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "validation" && impact.ItemCount == 1);
            Assert.Contains(plan.Impacts, impact =>
                impact.Category == "conditional-formatting" && impact.ItemCount == 1);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanIncludesQueryConnectionParameters() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            AttachCellBackedConnection(document, sheet, "A5");

            ExcelRowMutationPlan plan = sheet.PlanInsertRows(5);

            ExcelMutationImpact parameters = Assert.Single(
                plan.Impacts,
                impact => impact.Category == "connection-parameters");
            Assert.Equal(1, parameters.ItemCount);
        }

        [Fact]
        public void Test_StructuralRows_MutationPlanPreflightsUnmirroredPendingFormulas() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            for (int column = 1; column <= 128; column++) {
                sheet.CellValue(1, column, column);
            }
            sheet.CellFormula(1, 129, $"A{A1.MaxRows}");
            Assert.True(document.HasPendingDirectCellValues);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanInsertRows(1));

            Assert.Contains("row limit", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.True(document.HasPendingDirectCellValues);
        }
    }
}
