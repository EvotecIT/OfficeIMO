using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData("\n")]
    [InlineData("\r\n")]
    [InlineData("\u0085")]
    [InlineData("\u2028")]
    [InlineData("\u2029")]
    public async Task CurrencyContextDoesNotCrossLineBoundaries(string separator) {
        string source = "Total: 42 USD" + separator + "- Delivery included";
        var result = await new OfficeAiEngine(new Executor(Field("42", "e1"))).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields,
            Fields = new[] { new OfficeAiFieldDefinition("total", OfficeAiFieldType.Decimal) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Present, field.Status);
        Assert.Equal("42", field.NormalizedValue);
    }
}
