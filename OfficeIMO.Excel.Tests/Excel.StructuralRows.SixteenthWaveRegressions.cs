using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_IgnoresNonlocalConsolidationSourcesDuringPreflight() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            var otherSheetSource = new DataReference {
                Sheet = "Other",
                Reference = $"A{A1.MaxRows}"
            };
            var externalSource = new DataReference {
                Id = "rIdExternal",
                Sheet = "Data",
                Reference = $"A{A1.MaxRows}"
            };
            data.WorksheetPart.Worksheet.Append(
                new DataConsolidate(
                    new DataReferences(otherSheetSource, externalSource) { Count = 2U }) {
                    Function = DataConsolidateFunctionValues.Sum
                });

            data.InsertRows(1);

            Assert.Equal($"A{A1.MaxRows}", otherSheetSource.Reference!.Value);
            Assert.Equal($"A{A1.MaxRows}", externalSource.Reference!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RejectsDialogSheetVmlControls() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(2, 1).SetValue(1);
            DialogsheetPart dialogPart = document.WorkbookPartRoot.AddNewPart<DialogsheetPart>();
            dialogPart.DialogSheet = new DialogSheet();
            VmlDrawingPart vmlPart = dialogPart.AddNewPart<VmlDrawingPart>();
            XNamespace vml = "urn:schemas-microsoft-com:vml";
            XNamespace excel = "urn:schemas-microsoft-com:office:excel";
            var vmlDocument = new XDocument(
                new XElement(
                    "xml",
                    new XAttribute(XNamespace.Xmlns + "v", vml),
                    new XAttribute(XNamespace.Xmlns + "x", excel),
                    new XElement(
                        vml + "shape",
                        new XElement(
                            excel + "ClientData",
                            new XAttribute("ObjectType", "Button"),
                            new XElement(excel + "FmlaLink", "Data!$A$2")))));
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vmlDocument.Save(stream);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(2));

            Assert.Contains("form controls", exception.Message);
            Assert.Equal(1, sheet.CellAt(2, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_MaintainsCellSmartTagCollections() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            const string spreadsheetNamespace =
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var tags = new OpenXmlUnknownElement(string.Empty, "cellSmartTags", spreadsheetNamespace);
            tags.SetAttribute(new OpenXmlAttribute(string.Empty, "count", string.Empty, "2"));
            var deletedTag = new OpenXmlUnknownElement(string.Empty, "cellSmartTag", spreadsheetNamespace);
            deletedTag.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, "A2"));
            var survivingTag = new OpenXmlUnknownElement(string.Empty, "cellSmartTag", spreadsheetNamespace);
            survivingTag.SetAttribute(new OpenXmlAttribute(string.Empty, "r", string.Empty, "A5"));
            tags.Append(deletedTag, survivingTag);
            sheet.WorksheetPart.Worksheet.Append(tags);

            sheet.DeleteRows(2);

            Assert.Single(tags.ChildElements);
            Assert.Equal(
                "1",
                tags.GetAttributes().Single(attribute => attribute.LocalName == "count").Value);
            Assert.Equal(
                "A4",
                survivingTag.GetAttributes().Single(attribute => attribute.LocalName == "r").Value);

            sheet.DeleteRows(4);

            Assert.Null(tags.Parent);
        }

        [Fact]
        public void Test_StructuralRows_PreflightsVmlNoteAnchorOverflow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("keep");
            sheet.SetComment(1, 1, "Stationary note", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            XDocument vmlDocument;
            using (Stream stream = vmlPart.GetStream()) {
                vmlDocument = XDocument.Load(stream);
            }
            XNamespace excel = "urn:schemas-microsoft-com:office:excel";
            XElement anchor = vmlDocument.Descendants(excel + "Anchor").Single();
            string originalAnchor = $"1, 0, {A1.MaxRows - 1}, 0, 3, 0, {A1.MaxRows}, 0";
            anchor.Value = originalAnchor;
            using (Stream stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                vmlDocument.Save(stream);
            }

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(A1.MaxRows));

            Assert.Contains("note anchor", exception.Message);
            Assert.Equal("keep", sheet.CellAt(1, 1).GetValue<string>());
            using Stream unchangedStream = vmlPart.GetStream();
            Assert.Equal(
                originalAnchor,
                XDocument.Load(unchangedStream).Descendants(excel + "Anchor").Single().Value);
        }

        [Fact]
        public void Test_StructuralRows_RejectsDeletionOfCellBackedConnectionParameter() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            Parameter parameter = AttachCellBackedConnection(document, sheet, "A5");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(5));

            Assert.Contains("connection parameter", exception.Message);
            Assert.Equal("A5", parameter.Cell!.Value);
            Assert.Equal(1, sheet.CellAt(5, 1).GetValue<int>());
            Assert.Equal(
                1U,
                parameter.Ancestors<Parameters>().Single().Count!.Value);
        }

        [Fact]
        public void Test_StructuralRows_RejectsDeletionOfPivotSourceHeader() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = CreatePivotSheet(document);
            WorksheetSource source = Assert.Single(
                sheet.WorksheetPart.PivotTableParts).PivotTableCacheDefinitionPart!
                .PivotCacheDefinition!.CacheSource!.WorksheetSource!;

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(1));

            Assert.Contains("header row", exception.Message);
            Assert.Equal("A1:B3", source.Reference!.Value);
            Assert.Equal("Region", sheet.CellAt(1, 1).GetValue<string>());
        }
    }
}
