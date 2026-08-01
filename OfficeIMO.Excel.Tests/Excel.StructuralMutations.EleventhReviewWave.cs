using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ConnectionParameters_FollowTheirQualifiedWorksheetAcrossOwners() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet config = document.AddWorksheet("Config");
            Parameter parameter = AttachCellBackedConnection(document, data, "Config!A2");

            data.InsertRows(2);
            Assert.Equal("Config!A2", parameter.Cell!.Value);

            config.InsertRows(2);
            Assert.Equal("Config!A3", parameter.Cell!.Value);
            config.InsertColumns(1);
            Assert.Equal("Config!B3", parameter.Cell!.Value);
            config.MoveRange("B3", "C4");
            Assert.Equal("Config!C4", parameter.Cell!.Value);
        }

        [Fact]
        public void Test_RangeMutations_RejectPartialPivotCacheSources() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            PivotTableCacheDefinitionPart cachePart =
                document.WorkbookPartRoot.AddNewPart<PivotTableCacheDefinitionPart>();
            cachePart.PivotCacheDefinition = new PivotCacheDefinition(
                new CacheSource(new WorksheetSource { Sheet = "Data", Reference = "A1:B4" }) {
                    Type = SourceValues.Worksheet
                });

            InvalidOperationException move = Assert.Throws<InvalidOperationException>(() =>
                sheet.MoveRange("A1:A4", "C1"));
            Assert.Contains("pivot cache source", move.Message, StringComparison.OrdinalIgnoreCase);
            InvalidOperationException shift = Assert.Throws<InvalidOperationException>(() =>
                sheet.InsertCells("A2:A3", ExcelCellShiftDirection.Right));
            Assert.Contains("pivot cache source", shift.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_CellShifts_RejectWorkbookVmlControls() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            DialogsheetPart dialogPart = document.WorkbookPartRoot.AddNewPart<DialogsheetPart>();
            dialogPart.DialogSheet = new DialogSheet();
            VmlDrawingPart vmlPart = dialogPart.AddNewPart<VmlDrawingPart>();
            XNamespace vml = "urn:schemas-microsoft-com:vml";
            XNamespace excel = "urn:schemas-microsoft-com:office:excel";
            var vmlDocument = new XDocument(
                new XElement("xml",
                    new XAttribute(XNamespace.Xmlns + "v", vml),
                    new XAttribute(XNamespace.Xmlns + "x", excel),
                    new XElement(vml + "shape",
                        new XElement(excel + "ClientData",
                            new XAttribute("ObjectType", "Button"),
                            new XElement(excel + "FmlaLink", "Data!$A$2")))));
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) vmlDocument.Save(stream);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.InsertCells("A2", ExcelCellShiftDirection.Right));
            Assert.Contains("form controls", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Test_TableResize_RemapSortRangesAndRequestsFullCalculation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(1, 2, "B");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(2, 2, 2);
            sheet.CellFormula(1, 4, "SUM(Sales[A])");
            sheet.AddTable("A1:B2", true, "Sales", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            var sortState = new SortState(
                new SortCondition { Reference = "A2:A2" }) { Reference = "A2:B2" };
            table.Append(sortState);
            CalculationChainPart chainPart = document.WorkbookPartRoot.AddNewPart<CalculationChainPart>();
            chainPart.CalculationChain = new CalculationChain(
                new CalculationCell { CellReference = "A2", SheetId = 0 });

            sheet.ResizeTable("Sales", "A1:B5");

            Assert.Equal("A2:B5", sortState.Reference!.Value);
            Assert.Equal("A2:A5", Assert.Single(sortState.Elements<SortCondition>()).Reference!.Value);
            Assert.Null(document.WorkbookPartRoot.CalculationChainPart);
            CalculationProperties properties = document.WorkbookPartRoot.Workbook!
                .GetFirstChild<CalculationProperties>()!;
            Assert.True(properties.ForceFullCalculation?.Value ?? false);
            Assert.True(properties.FullCalculationOnLoad?.Value ?? false);
        }

        [Fact]
        public void Test_RangeMove_ReplacesDestinationHyperlinksAndRelationships() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetHyperlink(1, 1, "https://source.example/", display: "Source", style: false);
            sheet.SetHyperlink(1, 2, "https://destination.example/", display: "Destination", style: false);

            sheet.MoveRange("A1", "B1");

            Hyperlink hyperlink = Assert.Single(sheet.WorksheetPart.Worksheet!
                .GetFirstChild<Hyperlinks>()!.Elements<Hyperlink>());
            Assert.Equal("B1", hyperlink.Reference!.Value);
            Assert.Single(sheet.WorksheetPart.HyperlinkRelationships);
            Assert.Equal("https://source.example/", sheet.WorksheetPart.HyperlinkRelationships.Single().Uri.AbsoluteUri);
        }

        [Fact]
        public void Test_RangeMove_IgnoresAbsoluteAnchoredImages() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                byte[] imageBytes = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                using (var document = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = document.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Move");
                    sheet.AddImage(1, 1, imageBytes, "image/png", widthPixels: 16, heightPixels: 16);
                    document.Save();
                }
                using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                    Xdr.WorksheetDrawing drawing = package.WorkbookPart!.WorksheetParts.Single()
                        .DrawingsPart!.WorksheetDrawing!;
                    Xdr.OneCellAnchor oneCell = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
                    var absolute = new Xdr.AbsoluteAnchor(
                        new Xdr.Position { X = 0L, Y = 0L },
                        (Xdr.Extent)oneCell.Extent!.CloneNode(true),
                        (Xdr.Picture)oneCell.Descendants<Xdr.Picture>().Single().CloneNode(true),
                        new Xdr.ClientData());
                    oneCell.Remove();
                    drawing.Append(absolute);
                    drawing.Save();
                }
                using (var document = ExcelDocument.Load(path)) {
                    ExcelSheet sheet = document["Data"];
                    Assert.True(Assert.Single(sheet.Images).HasAbsoluteAnchor);
                    sheet.MoveRange("A1", "B1");
                    Assert.True(Assert.Single(sheet.Images).HasAbsoluteAnchor);
                }
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
