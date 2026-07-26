using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MovesButDoesNotResizeOneCellCommentAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(10, 2, "Move without sizing", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            SetCommentVmlPlacement(vmlPart, "1, 0, 3, 0, 4, 0, 6, 0", moveWithCells: true, sizeWithCells: false);

            sheet.InsertRows(6);

            Assert.Equal("1, 0, 3, 0, 4, 0, 6, 0", GetCommentVmlAnchor(vmlPart));

            sheet.InsertRows(4);

            Assert.Equal("1, 0, 4, 0, 4, 0, 7, 0", GetCommentVmlAnchor(vmlPart));
            Assert.True(sheet.HasComment(12, 2));
        }

        [Fact]
        public void Test_StructuralRows_PreservesAbsoluteCommentAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(5, 2, "Do not move or size", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            SetCommentVmlPlacement(vmlPart, "1, 0, 3, 0, 4, 0, 6, 0", moveWithCells: false, sizeWithCells: false);

            sheet.InsertRows(1, 2);

            Assert.Equal("1, 0, 3, 0, 4, 0, 6, 0", GetCommentVmlAnchor(vmlPart));
            Assert.True(sheet.HasComment(7, 2));
        }

        [Fact]
        public void Test_StructuralRows_AllowsInsertionThroughOneCellCommentAnchorAtGridBoundary() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(2, 2, "Tall note", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            string anchor = $"1, 0, 2, 0, 4, 0, {A1.MaxRows}, 0";
            SetCommentVmlPlacement(vmlPart, anchor, moveWithCells: true, sizeWithCells: false);

            sheet.InsertRows(5);

            Assert.Equal(anchor, GetCommentVmlAnchor(vmlPart));
        }

        [Fact]
        public void Test_StructuralRows_RejectsOneCellCommentAnchorOverflowAtomically() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(2, 2, "Boundary note", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            string anchor = $"1, 0, {A1.MaxRows - 2}, 0, 4, 0, {A1.MaxRows}, 0";
            SetCommentVmlPlacement(vmlPart, anchor, moveWithCells: true, sizeWithCells: false);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.InsertRows(A1.MaxRows - 1));

            Assert.Contains("comment note anchor", exception.Message);
            Assert.Equal(anchor, GetCommentVmlAnchor(vmlPart));
            Assert.True(sheet.HasComment(2, 2));
        }

        private static void SetCommentVmlPlacement(
            VmlDrawingPart vmlPart,
            string anchor,
            bool moveWithCells,
            bool sizeWithCells) {
            XDocument vml;
            using (Stream stream = vmlPart.GetStream()) {
                vml = XDocument.Load(stream);
            }
            XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
            XElement clientData = Assert.Single(vml.Descendants(excelNamespace + "ClientData"));
            clientData.SetElementValue(excelNamespace + "Anchor", anchor);
            clientData.Elements(excelNamespace + "MoveWithCells").Remove();
            clientData.Elements(excelNamespace + "SizeWithCells").Remove();
            if (moveWithCells) {
                clientData.AddFirst(new XElement(excelNamespace + "MoveWithCells"));
            }
            if (sizeWithCells) {
                clientData.AddFirst(new XElement(excelNamespace + "SizeWithCells"));
            }
            using Stream output = vmlPart.GetStream(FileMode.Create, FileAccess.Write);
            vml.Save(output);
        }

        private static string GetCommentVmlAnchor(VmlDrawingPart vmlPart) {
            using Stream stream = vmlPart.GetStream();
            XDocument vml = XDocument.Load(stream);
            XNamespace excelNamespace = "urn:schemas-microsoft-com:office:excel";
            return Assert.Single(vml.Descendants(excelNamespace + "Anchor")).Value;
        }
    }
}
