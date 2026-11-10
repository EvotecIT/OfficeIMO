using OfficeIMO.Data.Conversions;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// One row per entry of the design's loss table: the converter boundary emits the documented code, and FailOnLoss turns
// a loss into an exception while an informational finding does not.
public class LossReportingTests {
    public static TheoryData<string, string, string, string> Cases() => new() {
        { "# note\na: 1\n", "Yaml", "Json", OfficeDiagnosticCodes.CommentsOmitted },
        { "a: &x 1\nb: *x\n", "Yaml", "Json", OfficeDiagnosticCodes.YamlAliasExpanded },
        { "a: !Ref name\n", "Yaml", "Json", OfficeDiagnosticCodes.YamlTagDropped },
        { "? [a, b]\n: c\n", "Yaml", "Json", OfficeDiagnosticCodes.YamlComplexKey },
        { "200: ok\n", "Yaml", "Json", OfficeDiagnosticCodes.KeyTypeChanged },
        { "a: 1\n---\nb: 2\n", "Yaml", "Json", OfficeDiagnosticCodes.YamlDocumentsCombined },
        { "a: .nan\n", "Yaml", "Json", OfficeDiagnosticCodes.NonFiniteNumber },
        { "%YAML 1.1\n---\na: !!timestamp 2001-12-14\n", "Yaml", "Json", OfficeDiagnosticCodes.DatetimeAsString },
        { "{\"a\":1}", "Json", "Xml", OfficeDiagnosticCodes.ScalarTypeDropped },
        { "[1,2]", "Json", "Xml", OfficeDiagnosticCodes.RootWrapped },
        { "{\"r\":{\"a\":[1]}}", "Json", "Xml", OfficeDiagnosticCodes.XmlArrayShape },
        { "<r><a>1</a><b>2</b><a>3</a></r>", "Xml", "Json", OfficeDiagnosticCodes.XmlSiblingOrderChanged },
        { "<r><x>text <b>bold</b> more</x></r>", "Xml", "Json", OfficeDiagnosticCodes.XmlMixedContent },
        { "<?xml-stylesheet href=\"a.xsl\"?><r/>", "Xml", "Json", OfficeDiagnosticCodes.XmlProcessingInstructionDropped },
        { "[{\"a\":\"x\"},{\"a\":null},{\"a\":\"\"}]", "Json", "Csv", OfficeDiagnosticCodes.NullAsEmpty },
        { "[{\"id\":\"42\"}]", "Json", "Csv", OfficeDiagnosticCodes.ScalarTypeDropped },
        { "[{\"app.version\":\"1\"}]", "Json", "Csv", OfficeDiagnosticCodes.HeaderPathAmbiguous },
        { "[{\"tags\":[]}]", "Json", "Csv", OfficeDiagnosticCodes.EmptyCollectionDropped },
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void ReportsTheDocumentedCode(string input, string source, string target, string code) {
        (_, OfficeConversionReport report) = Convert.Run(input, Enum.Parse<DataFormat>(source), Enum.Parse<DataFormat>(target));
        Assert.Contains(code, Convert.Codes(report));
    }

    [Fact]
    public void FailOnLossRejectsLossButNotInformation() {
        var strict = new DataConversionOptions { FailOnLoss = true };
        Assert.Throws<OfficeConversionException>(() => Convert.Run("# note\na: 1\n", DataFormat.Yaml, DataFormat.Json, strict));
        (string json, OfficeConversionReport report) = Convert.Run("a: &x 1\nb: *x\n", DataFormat.Yaml, DataFormat.Json, strict);
        Assert.False(report.HasLoss);
        Assert.Contains("\"b\": 1", json);
    }

    [Fact]
    public void ArrayJoinIsReported() {
        var join = new DataConversionOptions { ArrayColumns = DataArrayColumns.Join };
        (string csv, OfficeConversionReport report) = Convert.Run("[{\"tags\":[\"a\",\"b\"]}]", DataFormat.Json, DataFormat.Csv, join);
        Assert.Contains(OfficeDiagnosticCodes.ArrayJoined, Convert.Codes(report));
        Assert.Equal("tags\n\"a, b\"\n", csv);
    }
}
