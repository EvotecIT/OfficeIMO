using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_RangeMove_RejectsPartialDestinationHyperlinkRanges() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Move");
            sheet.SetHyperlink(1, 2, "https://destination.example/", display: "Destination", style: false);
            Hyperlink hyperlink = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<Hyperlink>());
            hyperlink.Reference = "B1:D1";

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.MoveRange("A1", "B1"));

            Assert.Contains("partially overwrite hyperlink", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("B1:D1", hyperlink.Reference!.Value);
            Assert.Single(sheet.WorksheetPart.HyperlinkRelationships);
        }

        [Fact]
        public void Test_MutationRollback_RestoresDeletedHyperlinkRelationships() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetHyperlink(1, 1, "https://rollback.example/", display: "Rollback", style: false);
            HyperlinkRelationship relationship = Assert.Single(sheet.WorksheetPart.HyperlinkRelationships);
            string relationshipId = relationship.Id;
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.WorksheetPart.DeleteReferenceRelationship(relationship);
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            HyperlinkRelationship restored = Assert.Single(sheet.WorksheetPart.HyperlinkRelationships);
            Assert.Equal(relationshipId, restored.Id);
            Assert.Equal("https://rollback.example/", restored.Uri.AbsoluteUri);
            Assert.Equal(relationshipId, Assert.Single(sheet.WorksheetPart.Worksheet
                .Descendants<Hyperlink>()).Id!.Value);
        }

        [Fact]
        public void Test_ColumnCellAndMoveMutations_RemapViewScenarioAndWorkbookSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(2, 2, 1);
            data.CellValue(3, 3, 2);
            var pane = new Pane { TopLeftCell = "B2" };
            var selection = new Selection {
                ActiveCell = "B2",
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:C3" }
            };
            var view = new SheetView(pane, selection) {
                WorkbookViewId = 0U,
                TopLeftCell = "B2"
            };
            SheetData sheetData = data.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            data.WorksheetPart.Worksheet.InsertBefore(new SheetViews(view), sheetData);
            var input = new InputCells { CellReference = "B2", Val = "10" };
            data.WorksheetPart.Worksheet.Append(new Scenarios(
                new Scenario(input) { Name = "Case", Count = 1U }) {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "B2:C3" }
            });
            var localConsolidationSource = new DataReference {
                Sheet = "Data",
                Reference = "B2:C3"
            };
            data.WorksheetPart.Worksheet.Append(new DataConsolidate(
                new DataReferences(localConsolidationSource) { Count = 1U }) {
                Function = DataConsolidateFunctionValues.Sum
            });
            var consolidationSource = new DataReference {
                Sheet = "Data",
                Reference = "B2:C3"
            };
            summary.WorksheetPart.Worksheet.Append(new DataConsolidate(
                new DataReferences(consolidationSource) { Count = 1U }) {
                Function = DataConsolidateFunctionValues.Sum
            });
            var publishItem = new WebPublishItem {
                Id = 1U,
                DivId = "Data_1",
                SourceType = WebSourceValues.Range,
                SourceObject = "Data",
                SourceRef = "B2:C3",
                DestinationFile = "published.htm"
            };
            document.WorkbookPartRoot.Workbook!.Append(new WebPublishItems(publishItem) { Count = 1U });

            data.InsertColumns(1);
            Assert.Equal("C2", view.TopLeftCell!.Value);
            Assert.Equal("C2", pane.TopLeftCell!.Value);
            Assert.Equal("C2", selection.ActiveCell!.Value);
            Assert.Equal("C2:D3", selection.SequenceOfReferences!.InnerText);
            Assert.Equal("C2", input.CellReference!.Value);
            Assert.Equal("C2:D3", localConsolidationSource.Reference!.Value);
            Assert.Equal("C2:D3", consolidationSource.Reference!.Value);
            Assert.Equal("C2:D3", publishItem.SourceRef!.Value);

            data.InsertCells("C2:D3", ExcelCellShiftDirection.Right);
            Assert.Equal("E2", view.TopLeftCell!.Value);
            Assert.Equal("E2", input.CellReference!.Value);
            Assert.Equal("E2:F3", localConsolidationSource.Reference!.Value);
            Assert.Equal("E2:F3", consolidationSource.Reference!.Value);
            Assert.Equal("E2:F3", publishItem.SourceRef!.Value);

            data.MoveRange("E2:F3", "G4");
            Assert.Equal("G4", view.TopLeftCell!.Value);
            Assert.Equal("G4", pane.TopLeftCell!.Value);
            Assert.Equal("G4", selection.ActiveCell!.Value);
            Assert.Equal("G4:H5", selection.SequenceOfReferences!.InnerText);
            Assert.Equal("G4", input.CellReference!.Value);
            Assert.Equal("G4:H5", localConsolidationSource.Reference!.Value);
            Assert.Equal("G4:H5", consolidationSource.Reference!.Value);
            Assert.Equal("G4:H5", publishItem.SourceRef!.Value);
        }

        [Fact]
        public void Test_NamedStyleRedefinition_RejectsAmbiguousSharedStyleFormat() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Source");
            sheet.CellAt(1, 1).SetFillColor("C6EFCE");
            Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            CellStyle normal = stylesheet.CellStyles!.Elements<CellStyle>()
                .Single(style => style.Name?.Value == "Normal");
            uint normalFormatId = normal.FormatId!.Value;
            string normalFormatXml = stylesheet.CellStyleFormats!.Elements<CellFormat>()
                .ElementAt((int)normalFormatId).OuterXml;
            var alias = new CellStyle { Name = "Alias", FormatId = normalFormatId };
            stylesheet.CellStyles.Append(alias);
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.DefineNamedStyle("Alias", 1, 1));

            Assert.Contains("shares its base format", exception.Message);
            Assert.Equal(normalFormatId, normal.FormatId!.Value);
            Assert.Equal(normalFormatXml, stylesheet.CellStyleFormats.Elements<CellFormat>()
                .ElementAt((int)normalFormatId).OuterXml);
            Assert.Equal(normalFormatId, alias.FormatId!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
