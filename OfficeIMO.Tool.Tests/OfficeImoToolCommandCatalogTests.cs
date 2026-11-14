using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class OfficeImoToolCommandCatalogTests {
    [Fact]
    public async Task ConvertCatalogIncludesInputsHandledByIWorkAndReader() {
        await using var input = new MemoryStream();
        await using var output = new MemoryStream();
        using var error = new StringWriter();

        int exitCode = await OfficeImoToolApp.RunAsync(["help", "--format", "json"], input, output, error);

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        Assert.Equal(string.Empty, error.ToString());
        using JsonDocument catalog = JsonDocument.Parse(Encoding.UTF8.GetString(output.ToArray()));
        JsonElement convert = catalog.RootElement.GetProperty("groups").EnumerateArray()
            .Single(group => group.GetProperty("id").GetString() == "convert");
        var inputs = convert.GetProperty("inputFormats").EnumerateArray()
            .Select(format => format.GetString()!).ToHashSet(StringComparer.Ordinal);
        foreach (string format in new[] { "key", "numbers", "pages", "xls", "xlsb", "pst" }) {
            Assert.Contains(format, inputs);
        }
    }
}
