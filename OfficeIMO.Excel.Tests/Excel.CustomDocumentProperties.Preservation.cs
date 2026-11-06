using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RawCustomPropertiesSurviveEditsAndRetainedValuesRemainMutable(bool standardWriter) {
        using var input = new MemoryStream();
        using (var seed = ExcelDocument.Create(input)) {
            seed.AddSheet("Data").Cell(1, 1).SetValue("Before");
            seed.Save();
        }
        input.Position = 0;
        using (var package = SpreadsheetDocument.Open(input, true)) {
            var part = package.AddCustomFilePropertiesPart();
            const string properties = """
                <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
                  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                  xmlns:p="urn:preservation" p:root="retained">
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="12" name="LegacyText" linkTarget="NativeLink" p:marker="text"><vt:lpstr>café</vt:lpstr></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="20" name="ExactDecimal"><vt:decimal>1234567890.1234567890123456789</vt:decimal></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="31" name="Integer"><vt:i4>42</vt:i4></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="32" name="Short"><vt:i2>-3</vt:i2></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="33" name="Unsigned"><vt:ui8>18446744073709551615</vt:ui8></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="34" name="Reviewed"><vt:bool>1</vt:bool></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="40" name="Opaque" p:marker="opaque"><vt:vector size="1" baseType="lpstr"><vt:lpstr>carry</vt:lpstr></vt:vector></property>
                  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="41" name="Malformed"><vt:i4>not-an-integer</vt:i4></property>
                </Properties>
                """;
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(properties));
            part.FeedData(xml);
        }

        input.Position = 0;
        using var document = ExcelDocument.Load(input);
        ExcelCustomProperty text = document.CustomProperties["LegacyText"];
        ExcelCustomProperty number = document.CustomProperties["ExactDecimal"];
        ExcelCustomProperty opaque = document.CustomProperties["Opaque"];
        Assert.Equal("café", text.Text);
        Assert.Equal(1234567890.1234567890123456789M, number.Value);
        Assert.Equal(42, document.CustomProperties["Integer"].NumberInteger);
        Assert.Equal(-3, document.CustomProperties["Short"].NumberInteger);
        Assert.Equal(ulong.MaxValue, document.CustomProperties["Unsigned"].Value);
        Assert.True(document.CustomProperties["Reviewed"].Bool);
        Assert.False(document.CustomProperties.ContainsKey("Malformed"));
        document["Data"].Cell(1, 1).SetValue("After");
        document.SetCustomDocumentProperty("Added", true);
        var options = new ExcelSaveOptions { DisableFastPackageWriter = standardWriter };
        using var first = new MemoryStream();
        document.Save(first, options);
        Assert.Same(text, document.CustomProperties["LegacyText"]);
        Assert.Same(opaque, document.CustomProperties["Opaque"]);

        XNamespace cp = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        XNamespace vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        XNamespace metadata = "urn:preservation";
        using (var package = SpreadsheetDocument.Open(new MemoryStream(first.ToArray()), false)) {
            using var stream = package.CustomFilePropertiesPart!.GetStream();
            XElement root = XDocument.Load(stream).Root!;
            var properties = root.Elements(cp + "property").ToDictionary(p => (string)p.Attribute("name")!);
            Assert.Equal("retained", (string?)root.Attribute(metadata + "root"));
            Assert.Equal("café", (string?)properties["LegacyText"].Element(vt + "lpstr"));
            Assert.Equal("12", (string?)properties["LegacyText"].Attribute("pid"));
            Assert.Equal("NativeLink", (string?)properties["LegacyText"].Attribute("linkTarget"));
            Assert.Equal("text", (string?)properties["LegacyText"].Attribute(metadata + "marker"));
            Assert.Equal("1234567890.1234567890123456789", (string?)properties["ExactDecimal"].Element(vt + "decimal"));
            Assert.Equal("20", (string?)properties["ExactDecimal"].Attribute("pid"));
            Assert.Equal("42", (string?)properties["Integer"].Element(vt + "i4"));
            Assert.Equal("-3", (string?)properties["Short"].Element(vt + "i2"));
            Assert.Equal("18446744073709551615", (string?)properties["Unsigned"].Element(vt + "ui8"));
            Assert.Equal("1", (string?)properties["Reviewed"].Element(vt + "bool"));
            Assert.Equal("carry", (string?)properties["Opaque"].Element(vt + "vector")!.Element(vt + "lpstr"));
            Assert.Equal("1", (string?)properties["Opaque"].Element(vt + "vector")!.Attribute("size"));
            Assert.Equal("opaque", (string?)properties["Opaque"].Attribute(metadata + "marker"));
            Assert.Equal("not-an-integer", (string?)properties["Malformed"].Element(vt + "i4"));
            Assert.Equal(properties.Count, properties.Values.Select(p => (string?)p.Attribute("pid")).Distinct().Count());
        }

        text.Value = "Revised";
        number.Value = 12.3400M;
        opaque.Value = false;
        opaque.PropertyType = ExcelCustomPropertyType.YesNo;
        document.SetCustomDocumentProperty("Short", 8);
        Assert.True(document.RemoveCustomDocumentProperty("Integer"));
        using var second = new MemoryStream();
        document.Save(second, options);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(second.ToArray()), false)) {
            using var stream = package.CustomFilePropertiesPart!.GetStream();
            var properties = XDocument.Load(stream).Root!.Elements(cp + "property").ToDictionary(p => (string)p.Attribute("name")!);
            Assert.Equal("Revised", (string?)properties["LegacyText"].Element(vt + "lpwstr"));
            Assert.Null(properties["LegacyText"].Element(vt + "lpstr"));
            Assert.Equal("NativeLink", (string?)properties["LegacyText"].Attribute("linkTarget"));
            Assert.Equal("12", (string?)properties["LegacyText"].Attribute("pid"));
            Assert.Equal("12.3400", (string?)properties["ExactDecimal"].Element(vt + "decimal"));
            Assert.Equal("false", (string?)properties["Opaque"].Element(vt + "bool"));
            Assert.Null(properties["Opaque"].Element(vt + "vector"));
            Assert.Equal("opaque", (string?)properties["Opaque"].Attribute(metadata + "marker"));
            Assert.False(properties.ContainsKey("Integer"));
            Assert.Equal("not-an-integer", (string?)properties["Malformed"].Element(vt + "i4"));
            Assert.Equal("8", (string?)properties["Short"].Element(vt + "i4"));
            Assert.Equal("32", (string?)properties["Short"].Attribute("pid"));
        }
    }
}
