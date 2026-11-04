using System.IO.Compression;
using System.Text;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData("reader", false)]
    [InlineData("reader", true)]
    [InlineData("snapshot", false)]
    [InlineData("snapshot", true)]
    [InlineData("loaded", false)]
    [InlineData("loaded", true)]
    [InlineData("parallel", false)]
    [InlineData("parallel", true)]
    [InlineData("grid", false)]
    [InlineData("grid", true)]
    [InlineData("grid-readonly", false)]
    [InlineData("grid-loaded", true)]
    public void Reader_OmittedReferencesFollowDocumentOrder(string route, bool utf16) {
        using var package = new MemoryStream();
        using (var zip = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true)) {
            Add("[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
                """);
            Add("_rels/.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
                """);
            Add("xl/workbook.xml", """
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>
                """);
            Add("xl/_rels/workbook.xml.rels", """
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>
                """);
            Add("xl/worksheets/sheet1.xml", """
                <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row><c t="str"><v>Id</v></c><c t="str"><v>Name</v></c><c t="str"><v>Amount</v></c></row><row><c><v>1</v></c><c t="str"><v>Alpha</v></c><c><v>7</v></c></row><row><c><v>2</v></c><c/><c><v>9</v></c></row><row r="4"><c r="A4"><v>3</v></c><c t="str"><v>Gamma</v></c><c><v>11</v></c></row><row><c><v>4</v></c><c t="str"><v>Delta</v></c><c><v>13</v></c></row></sheetData></worksheet>
                """, utf16 ? Encoding.Unicode : Encoding.UTF8);

            void Add(string name, string xml, Encoding? encoding = null) {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), encoding ?? new UTF8Encoding(false));
                writer.Write(xml);
            }
        }
        byte[] bytes = package.ToArray();
        OmittedReferenceRecord[] rows;
        if (route == "reader") {
            using var reader = ExcelDocument.OpenReader(bytes);
            rows = reader.RowsAs<OmittedReferenceRecord>().ToArray();
        } else {
            using var input = new MemoryStream(bytes);
            using var document = ExcelDocument.Load(input, new ExcelLoadOptions {
                AccessMode = route == "grid-readonly" ? DocumentAccessMode.ReadOnly : DocumentAccessMode.ReadWrite
            });
            string? originalXml = null;
            if (route is "loaded" or "parallel" or "grid-loaded") {
                var worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet;
                if (route == "grid-loaded") originalXml = worksheet.OuterXml;
            }
            var sheet = document.Sheets.Single();
            if (route.StartsWith("grid", StringComparison.Ordinal)) {
                // Forward and backward reads exercise the row and cell caches too.
                foreach (int row in new[] { 2, 3, 4, 5, 2 }) {
                    Assert.True(sheet.TryGetCellValueSnapshot(row, 1, out var id));
                    Assert.Equal((row - 1).ToString(System.Globalization.CultureInfo.InvariantCulture), id!.RawValue);
                    Assert.True(sheet.TryGetCellValueSnapshot(row, 3, out var amount));
                    Assert.Equal((row * 2 + 3).ToString(System.Globalization.CultureInfo.InvariantCulture), amount!.RawValue);
                }
            }
            if (originalXml != null) Assert.Equal(originalXml, document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet.OuterXml);
            rows = route == "parallel"
                ? sheet.RowsAsParallel<OmittedReferenceRecord>().ToArray()
                : sheet.RowsAs<OmittedReferenceRecord>().ToArray();
        }
        Assert.Equal(new[] { 1, 2, 3, 4 }, rows.Select(row => row.Id));
        Assert.Equal(new string?[] { "Alpha", null, "Gamma", "Delta" }, rows.Select(row => row.Name));
        Assert.Equal(new[] { 7d, 9d, 11d, 13d }, rows.Select(row => row.Amount));
    }

    public sealed class OmittedReferenceRecord {
        public int Id { get; set; }
        public string? Name { get; set; }
        public double Amount { get; set; }
    }
}
