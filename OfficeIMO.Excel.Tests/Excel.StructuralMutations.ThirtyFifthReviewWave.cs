using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_InCellImage_RejectsInvalidAlternativeTextBeforeMutation() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Images");
            int workbookPartCount = document.WorkbookPartRoot.Parts.Count();

            Assert.Throws<ArgumentException>(() =>
                sheet.SetInCellImage(1, 1, TinyPng, altText: "Bad\u0001Text"));

            Assert.Equal(workbookPartCount, document.WorkbookPartRoot.Parts.Count());
            Assert.Null(document.WorkbookPartRoot.CellMetadataPart);
            Assert.Empty(document.WorkbookPartRoot.RdRichValueParts);
            Assert.Empty(sheet.GetInCellImages());
        }

        [Fact]
        public void Test_QueryInspection_BoundsImportedConnectionMetadata() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Data");
            ExtendedPart part = document.WorkbookPartRoot.AddExtendedPart(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml",
                "xml");
            byte[] payload = Encoding.UTF8.GetBytes(
                "<connections>" + new string('x', ExcelDocument.MaximumWorkbookConnectionMetadataCharacters) + "</connections>");
            using (var stream = new MemoryStream(payload, writable: false)) {
                part.FeedData(stream);
            }

            Assert.Empty(document.GetQueryBackedTables());
            ExcelFeatureReport report = document.InspectFeatures();
            Assert.NotNull(report);
        }

        [Fact]
        public void Test_QueryBackedTable_ReusesPreservedConnectionPartLifecycle() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("Data");
            OpenXmlPart preserved = document.AddWorkbookConnectionMetadata(
                "<connections xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\">"
                + "<connection id=\"77\" name=\"Existing\" type=\"1\"/>"
                + "</connections>");

            ExcelQueryBackedTableInfo query = document.AddQueryBackedTable(new ExcelQueryBackedTableOptions {
                ConnectionName = "Orders",
                WorksheetName = "Data",
                TableName = "OrderResults",
                ColumnNames = new[] { "Id" }
            });

            Assert.Null(document.WorkbookPartRoot.ConnectionsPart);
            OpenXmlPart connectionPart = Assert.Single(
                document.WorkbookPartRoot.Parts.Select(pair => pair.OpenXmlPart),
                part => part.RelationshipType.EndsWith("/connections", StringComparison.Ordinal));
            Assert.Same(preserved, connectionPart);
            XDocument authored = LoadPartXml(connectionPart);
            Assert.Equal(2, authored.Descendants().Count(element => element.Name.LocalName == "connection"));
            Assert.Contains(document.GetQueryBackedTables(), item => item.ConnectionId == query.ConnectionId);

            Assert.True(document.RemoveQueryBackedTable(query.TableName));
            XDocument detached = LoadPartXml(connectionPart);
            Assert.Single(detached.Descendants(), element => element.Name.LocalName == "connection");
            Assert.Contains(detached.Descendants(), element => element.Name.LocalName == "connection"
                && element.Attribute("id")?.Value == "77");
        }

        private static XDocument LoadPartXml(OpenXmlPart part) {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return XDocument.Load(stream);
        }
    }
}
