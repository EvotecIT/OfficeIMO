using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static void ValidateSpreadsheetDocument(string filePath, SpreadsheetDocument spreadsheet) {
            Assert.NotNull(spreadsheet.WorkbookPart);
            Assert.True(spreadsheet.WorkbookPart!.WorksheetParts.Any());

            using Package package = Package.Open(filePath, FileMode.Open, FileAccess.Read);
            foreach (PackagePart part in package.GetParts().Where(p => p.ContentType == "application/xml")) {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                XDocument.Load(stream);
            }
        }
    }
}
