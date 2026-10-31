using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeThemeDocumentSource {
    [Theory]
    [InlineData("docx", "word/document.xml")]
    [InlineData("xlsx", "xl/workbook.xml")]
    [InlineData("pptx", "ppt/presentation.xml")]
    public void ThemeSourceReadsReferencedPaletteAndFontsFromEachPackageFamily(string extension, string main) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
        try {
            WriteSource(path, main);
            OfficeTheme theme = OfficeTheme.FromDocument(path);
            Assert.Equal("Corporate", theme.Name);
            Assert.Equal(OfficeColor.ParseHex("123456"), theme.Palette.Accent1);
            Assert.Equal(OfficeColor.ParseHex("101112"), theme.Palette.Text);
            Assert.Equal("Georgia", theme.Fonts.Heading.Family);
            Assert.Equal("Verdana", theme.Fonts.Body.Family);
            Assert.Equal("Verdana", theme.ChartStyle.AxisFont.Family);
            Assert.Equal(OfficeTheme.Modern.Page.MarginLeft, theme.Page.MarginLeft);
            Assert.Equal(OfficeTheme.Modern.TypeScale.Body, theme.TypeScale.Body);
            Assert.Equal(OfficeColor.ParseHex("3F64D7"), OfficeTheme.Modern.Palette.Accent1);
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ThemeSourceRejectsExternalRelationshipsAndDocumentTypeDeclarations(bool external) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            WriteSource(path, "word/document.xml", external);
            if (!external) {
                using var package = ZipFile.Open(path, ZipArchiveMode.Update);
                package.GetEntry("assets/company.xml")!.Delete();
                Add(package, "assets/company.xml", "<!DOCTYPE theme [<!ENTITY x 'no'>]><theme>&x;</theme>");
            }
            if (external) Assert.Throws<InvalidDataException>(() => OfficeTheme.FromDocument(path));
            else Assert.Throws<System.Xml.XmlException>(() => OfficeTheme.FromDocument(path));
        } finally { File.Delete(path); }
    }

    private static void WriteSource(string path, string main, bool external = false) {
        using var package = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(package, "_rels/.rels", Relationships("officeDocument", main));
        if (main.StartsWith("ppt/", StringComparison.Ordinal)) {
            Add(package, main, "<p:presentation xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><p:sldMasterIdLst><p:sldMasterId id='2147483648' r:id='master'/></p:sldMasterIdLst></p:presentation>");
            Add(package, "ppt/_rels/presentation.xml.rels", Relationships("slideMaster", "slideMasters/master.xml", id: "master"));
            Add(package, "ppt/slideMasters/_rels/master.xml.rels", Relationships("theme", "../../assets/company.xml"));
        } else {
            string folder = main.Substring(0, main.IndexOf('/'));
            Add(package, folder + "/_rels/" + Path.GetFileName(main) + ".rels", Relationships("theme", external ? "https://invalid.example/theme.xml" : "../assets/company.xml", external));
        }
        Add(package, "assets/company.xml", """
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Corporate">
              <a:themeElements><a:clrScheme name="Corporate">
                <a:dk1><a:sysClr val="windowText" lastClr="101112"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
                <a:dk2><a:srgbClr val="444444"/></a:dk2>
                <a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2>
                <a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="444444"/></a:accent4>
                <a:accent5><a:srgbClr val="555555"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6>
              </a:clrScheme><a:fontScheme name="Corporate"><a:majorFont><a:latin typeface="Georgia"/></a:majorFont>
                <a:minorFont><a:latin typeface="Verdana"/></a:minorFont></a:fontScheme></a:themeElements>
            </a:theme>
            """);
        Add(package, "theme/theme1.xml", "<decoy/>");
    }

    private static string Relationships(string type, string target, bool external = false, string id = "source") =>
        "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='" + id +
        "' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + type + "' Target='" + target +
        "'" + (external ? " TargetMode='External'" : string.Empty) + "/></Relationships>";

    private static void Add(ZipArchive package, string part, string text) {
        using Stream output = package.CreateEntry(part).Open();
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }
}
