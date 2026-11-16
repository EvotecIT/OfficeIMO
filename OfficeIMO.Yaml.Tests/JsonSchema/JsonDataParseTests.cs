using OfficeIMO.Json;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// JsonData.Parse against JSONTestSuite (OfficeIMO.TestAssets/Json/JSONTestSuite): accept cases parse, reject cases
// throw FormatException, and implementation-defined cases neither crash nor hang.
public class JsonDataParseTests {
    private static string Directory => Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Json", "JSONTestSuite", "test_parsing");

    public static TheoryData<string> Cases() {
        var data = new TheoryData<string>();
        foreach (string path in System.IO.Directory.GetFiles(Directory, "*.json").OrderBy(p => p, StringComparer.Ordinal)) data.Add(Path.GetFileName(path));
        return data;
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void AcceptsValidAndRejectsInvalidJson(string name) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(Directory, name));
        string text;
        try {
            text = new System.Text.UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
        } catch (System.Text.DecoderFallbackException) {
            // Invalid UTF-8 never reaches the parser, which reads decoded text.
            Assert.False(name.StartsWith("y_", StringComparison.Ordinal), name);
            return;
        }
        Exception? failure = Record.Exception(() => JsonData.Parse(text));
        if (failure is not null) Assert.IsType<FormatException>(failure);
        if (name.StartsWith("y_", StringComparison.Ordinal)) Assert.Null(failure);
        if (name.StartsWith("n_", StringComparison.Ordinal)) Assert.NotNull(failure);
    }

    [Fact]
    public void NumbersStayExactBeyondDoublePrecision() {
        JsonData big = JsonData.Parse("[12345678901234567890123456789012345678901, 0.1, 1e400, 1.0]");

        Assert.Equal("[12345678901234567890123456789012345678901,0.1,1e+400,1]", big.ToString());
        Assert.True(big.Items[0].IsInteger);
        Assert.True(big.Items[3].IsInteger);
        Assert.True(big.Items[1].TryGetDecimal(out decimal tenth));
        Assert.Equal(0.1m, tenth);
        Assert.False(big.Items[2].TryGetDecimal(out _));
        Assert.True(JsonData.DeepEquals(JsonData.Parse("{\"a\":[1.50]}"), JsonData.Parse("{ \"a\" : [ 15e-1 ] }")));
    }

    [Fact]
    public void ValuesRecordTheirLineAndColumn() {
        JsonData data = JsonData.Parse("{\n  \"a\": [1,\n    true]\n}");

        JsonData flag = data.Properties[0].Value.Items[1];
        Assert.Equal((3, 5), (flag.Line, flag.Column));
        FormatException error = Assert.Throws<FormatException>(() => JsonData.Parse("{\n  \"a\": tru }"));
        Assert.Contains("line 2", error.Message, StringComparison.Ordinal);
    }
}
