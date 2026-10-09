using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioShapeDataSchemaPreservationTests {
    private static readonly XNamespace Legacy = "http://schemas.microsoft.com/visio/2003/core";
    private static readonly XNamespace Modern = "http://schemas.microsoft.com/office/visio/2012/main";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RetainedValuesKeepProducerFormulasErrorsAndNullMarkers(bool connector) {
        VisioDocument document = Load(connector);
        VisioShapeDataRow cached = Row(document, connector, "Cached");
        VisioShapeDataRow nullable = Row(document, connector, "Nullable");
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
            .Field("Cached", "Updated label", VisioShapeDataType.String, "schema default", "Updated prompt", "@", sortKey: "020", verify: true)
            .Field("Nullable", defaultValue: "schema default");

        Apply(document, connector, schema);

        Assert.Same(cached, Row(document, connector, "Cached"));
        Assert.Equal("computed", cached.Value);
        Assert.Equal("User.Producer", cached.ValueFormula);
        Assert.Equal(VisioShapeDataType.String, cached.Type);
        Assert.Equal(string.Empty, nullable.Value);
        Assert.Equal("Inh", nullable.ValueFormula);
        Assert.Equal("computed", Data(document, connector)["Cached"]);
        Assert.Equal(string.Empty, Data(document, connector)["Nullable"]);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", "computed", null, "User.Producer", "#REF!");
            AssertCell(candidate, "Nullable", string.Empty, "null", "Inh", "producer-specific error");
            XElement metadata = Prop(candidate, "Cached");
            Assert.Equal("Updated label", metadata.Element(Legacy + "Label")!.Value);
            Assert.Equal("Updated prompt", metadata.Element(Legacy + "Prompt")!.Value);
            Assert.Equal("0", metadata.Element(Legacy + "Type")!.Value);
            Assert.Equal("@", metadata.Element(Legacy + "Format")!.Value);
            Assert.Equal("020", metadata.Element(Legacy + "SortKey")!.Value);
            Assert.Equal("1", metadata.Element(Legacy + "Verify")!.Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitDefaultsReplaceSameEmptyNullAndOtherExistingValues(bool connector) {
        VisioDocument document = Load(connector);
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
            .Field("Cached", defaultValue: "replacement")
            .Field("Nullable", defaultValue: string.Empty);

        Apply(document, connector, schema, overwriteValues: true);

        Assert.Equal("replacement", Row(document, connector, "Cached").Value);
        Assert.Equal(string.Empty, Row(document, connector, "Nullable").Value);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", "replacement", null, null, null);
            AssertCell(candidate, "Nullable", string.Empty, null, null, null);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RetainedDictionaryOverrideDoesNotReplaceTheProducerRow(bool connector) {
        VisioDocument document = Load(connector);
        VisioShapeDataRow row = Row(document, connector, "Cached");
        Data(document, connector)["Cached"] = "dictionary override";
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
            .Field("Cached", "Updated label", defaultValue: "schema default");

        Apply(document, connector, schema);

        Assert.Equal("computed", row.Value);
        Assert.Equal("User.Producer", row.ValueFormula);
        Assert.Equal("dictionary override", Data(document, connector)["Cached"]);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", "dictionary override", null, null, null);
        }

        Data(document, connector).Remove("Cached");
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", "computed", null, "User.Producer", "#REF!");
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultsPopulateAbsentValuesAndMaterializeDictionaryOnlyFields(bool connector) {
        VisioDocument document = Load(connector);
        Data(document, connector)["DictionaryOnly"] = "retained dictionary";
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
            .Field("MissingValue", "Present row", defaultValue: "value default")
            .Field("MissingRow", "New row", defaultValue: "row default")
            .Field("DictionaryOnly", "Dictionary row", defaultValue: "unused default");

        Apply(document, connector, schema);

        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "MissingValue", "value default", null, null, null, unit: null);
            AssertCell(candidate, "MissingRow", "row default", null, null, null, unit: null);
            AssertCell(candidate, "DictionaryOnly", "retained dictionary", null, null, null, unit: null);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PresentValueCellsWithoutCachedResultsAreRetainedUnlessOverwritten(bool connector) {
        using var stream = new MemoryStream();
        byte[] original = Load(connector).ToBytes();
        stream.Write(original, 0, original.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("visio/pages/page1.xml")!;
            XDocument xml;
            using (Stream input = entry.Open()) xml = XDocument.Load(input);
            XElement section = xml.Descendants(Modern + "Section").Single(element =>
                (string?)element.Attribute("N") is "Prop" or "Property");
            XElement value = section.Elements(Modern + "Row").Single(row => (string?)row.Attribute("N") == "Cached")
                .Elements(Modern + "Cell").Single(cell => (string?)cell.Attribute("N") == "Value");
            value.Attribute("V")!.Remove();
            value.SetAttributeValue("F", "1/0");
            value.SetAttributeValue("E", "#DIV/0!");
            section.Add(new XElement(Modern + "Row", new XAttribute("N", "ErrorOnly"), new XAttribute("IX", "3"),
                new XElement(Modern + "Cell", new XAttribute("N", "Value"), new XAttribute("E", "#REF!"))));
            section.Add(new XElement(Modern + "Row", new XAttribute("N", "EmptyValue"), new XAttribute("IX", "4"),
                new XElement(Modern + "Cell", new XAttribute("N", "Value"))));
            entry.Delete();
            using Stream output = archive.CreateEntry("visio/pages/page1.xml").Open();
            xml.Save(output);
        }
        stream.Position = 0;
        VisioDocument document = VisioDocument.Load(stream);
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create()
            .Field("Cached", "Updated label", defaultValue: "formula replacement")
            .Field("ErrorOnly", defaultValue: "error replacement")
            .Field("EmptyValue", defaultValue: "empty default")
            .Field("MissingValue", defaultValue: "missing default");
        Assert.Null(Row(document, connector, "Cached").Value);
        Assert.Null(Row(document, connector, "ErrorOnly").Value);

        Apply(document, connector, schema);

        Assert.Null(Row(document, connector, "Cached").Value);
        Assert.Null(Row(document, connector, "ErrorOnly").Value);
        foreach (VisioDocument candidate in new[] { document, VisioDocument.Load(new MemoryStream(document.ToBytes())) }) {
            AssertUncachedCell(candidate, "Cached", "1/0", "#DIV/0!");
            AssertUncachedCell(candidate, "ErrorOnly", null, "#REF!");
        }
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", string.Empty, null, "1/0", "#DIV/0!");
            AssertCell(candidate, "ErrorOnly", string.Empty, null, null, "#REF!", unit: null);
            AssertCell(candidate, "EmptyValue", "empty default", null, null, null, unit: null);
            AssertCell(candidate, "MissingValue", "missing default", null, null, null, unit: null);
        }

        Apply(document, connector, schema, overwriteValues: true);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "Cached", "formula replacement", null, null, null);
            AssertCell(candidate, "ErrorOnly", "error replacement", null, null, null, unit: null);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AuthoredFormulaWithoutCachedValueIsRetainedUnlessOverwritten(bool connector) {
        VisioDocument document = Load(connector);
        VisioShapeDataRow row = new("FormulaOnly") { ValueFormula = "GUARD(\"uncomputed\")" };
        if (connector) document.Pages[0].Connectors.Single().ShapeData.Add(row);
        else document.Pages[0].Shapes.Single().ShapeData.Add(row);
        VisioShapeDataSchema schema = VisioShapeDataSchema.Create().Field("FormulaOnly", defaultValue: "replacement");

        Apply(document, connector, schema);

        Assert.Null(row.Value);
        Assert.Equal("GUARD(\"uncomputed\")", row.ValueFormula);
        AssertUncachedCell(document, "FormulaOnly", "GUARD(\"uncomputed\")", null);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "FormulaOnly", string.Empty, null, "GUARD(\"uncomputed\")", null, unit: null);
        }

        Apply(document, connector, schema, overwriteValues: true);
        foreach (VisioDocument candidate in Candidates(document)) {
            AssertCell(candidate, "FormulaOnly", "replacement", null, null, null, unit: null);
        }
    }

    private static void AssertUncachedCell(VisioDocument document, string name, string? formula, string? error) {
        using var archive = new ZipArchive(new MemoryStream(document.ToBytes()), ZipArchiveMode.Read);
        using Stream input = archive.GetEntry("visio/pages/page1.xml")!.Open();
        XElement cell = XDocument.Load(input).Descendants(Modern + "Row").Single(row => (string?)row.Attribute("N") == name)
            .Elements(Modern + "Cell").Single(element => (string?)element.Attribute("N") == "Value");
        Assert.Null(cell.Attribute("V"));
        Assert.Equal(formula, (string?)cell.Attribute("F"));
        Assert.Equal(error, (string?)cell.Attribute("E"));
    }

    private static VisioDocument Load(bool connector) {
        string endpoints = connector ? "<XForm1D><BeginX>1</BeginX><BeginY>1</BeginY><EndX>3</EndX><EndY>1</EndY></XForm1D>" : string.Empty;
        string source = $"<VisioDocument xmlns='{Legacy}'><Pages><Page ID='0'><Shapes><Shape ID='1'>"
            + "<XForm><Width>2</Width><Height>1</Height></XForm>" + endpoints
            + "<Prop ID='0' NameU='Cached'><Value Unit='STR' F='User.Producer' Err='#REF!'>computed</Value><Label>Original label</Label></Prop>"
            + "<Prop ID='1' NameU='Nullable'><Value V='null' Unit='STR' F='Inh' Err='producer-specific error'/></Prop>"
            + "<Prop ID='2' NameU='MissingValue'><Label>No value</Label></Prop>"
            + "</Shape></Shapes></Page></Pages></VisioDocument>";
        return VisioDocument.LoadLegacyXml(new MemoryStream(Encoding.UTF8.GetBytes(source))).Value;
    }

    private static VisioShapeDataRow Row(VisioDocument document, bool connector, string name) => connector
        ? document.Pages[0].Connectors.Single().FindShapeData(name)!
        : document.Pages[0].Shapes.Single().FindShapeData(name)!;

    private static IDictionary<string, string> Data(VisioDocument document, bool connector) => connector
        ? document.Pages[0].Connectors.Single().Data
        : document.Pages[0].Shapes.Single().Data;

    private static void Apply(VisioDocument document, bool connector, VisioShapeDataSchema schema, bool overwriteValues = false) {
        if (connector) schema.ApplyTo(document.Pages[0].Connectors.Single(), overwriteValues);
        else schema.ApplyTo(document.Pages[0].Shapes.Single(), overwriteValues);
    }

    private static IEnumerable<VisioDocument> Candidates(VisioDocument document) {
        yield return document;
        yield return VisioDocument.Load(new MemoryStream(document.ToBytes()));
        yield return VisioDocument.LoadLegacyXml(new MemoryStream(document.ToLegacyXmlResult().Value)).Value;
    }

    private static XElement Prop(VisioDocument document, string name) => XDocument.Load(new MemoryStream(document.ToLegacyXmlResult().Value))
        .Descendants(Legacy + "Prop").Single(row => (string?)row.Attribute("NameU") == name);

    private static void AssertCell(VisioDocument document, string name, string value, string? marker, string? formula, string? error, string? unit = "STR") {
        XElement cell = Prop(document, name).Element(Legacy + "Value")!;
        Assert.Equal(value, cell.Value);
        Assert.Equal(marker, (string?)cell.Attribute("V"));
        Assert.Equal(formula, (string?)cell.Attribute("F"));
        Assert.Equal(error, (string?)cell.Attribute("Err"));
        Assert.Equal(unit, (string?)cell.Attribute("Unit"));
    }
}
