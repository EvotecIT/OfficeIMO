using System.IO.Compression;
using System.Security;

namespace OfficeIMO.Odbc.Tests;

/// <summary>Writes independent minimal OOXML input rather than using the engine under test as its oracle.</summary>
internal sealed class WorkbookFixture : IDisposable {
    internal string DirectoryPath { get; } = Path.Combine(Path.GetTempPath(), "OfficeIMO.Odbc.Tests", Guid.NewGuid().ToString("N"));
    internal string FilePath => Path.Combine(DirectoryPath, "customers;2026.xlsx");

    internal WorkbookFixture() { Directory.CreateDirectory(DirectoryPath); Write(); }

    internal void Write(string alice = "Alicja Żółć") {
        using var archive = ZipFile.Open(FilePath, File.Exists(FilePath) ? ZipArchiveMode.Update : ZipArchiveMode.Create);
        Put(archive, "[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>
            """);
        Put(archive, "_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>
            """);
        Put(archive, "xl/workbook.xml", """
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Customers" sheetId="1" r:id="rId1"/></sheets></workbook>
            """);
        Put(archive, "xl/_rels/workbook.xml.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>
            """);
        Put(archive, "xl/worksheets/sheet1.xml", $$"""
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:E3"/><sheetData>
            <row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Country</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c><c r="D1" t="inlineStr"><is><t>Mixed</t></is></c><c r="E1" t="inlineStr"><is><t>Optional</t></is></c></row>
            <row r="2"><c r="A2" t="inlineStr"><is><t>{{SecurityElement.Escape(alice)}}</t></is></c><c r="B2" t="inlineStr"><is><t>Poland</t></is></c><c r="C2"><v>12.5</v></c><c r="D2"><v>7</v></c><c r="E2" t="inlineStr"><is><t></t></is></c></row>
            <row r="3"><c r="A3" t="inlineStr"><is><t>Bob</t></is></c><c r="B3" t="inlineStr"><is><t>Germany</t></is></c><c r="C3"><v>20</v></c><c r="D3" t="inlineStr"><is><t>001</t></is></c></row>
            </sheetData></worksheet>
            """);
    }

    private static void Put(ZipArchive archive, string name, string xml) {
        if (archive.Mode == ZipArchiveMode.Update) archive.GetEntry(name)?.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(xml);
    }

    public void Dispose() => Directory.Delete(DirectoryPath, true);
}
