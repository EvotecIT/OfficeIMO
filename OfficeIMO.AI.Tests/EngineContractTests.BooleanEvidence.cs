using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData("untrue", "true", false)]
    [InlineData("falsehood", "false", false)]
    [InlineData("true_value", "true", false)]
    [InlineData("true9", "true", false)]
    [InlineData("étrue", "true", false)]
    [InlineData("true\u0301", "true", false)]
    [InlineData("\U00010400true", "true", false)]
    [InlineData("true\U00010400", "true", false)]
    [InlineData("Value: true.", "true", true)]
    [InlineData("Value: FALSE;", "FALSE", true)]
    [InlineData("Value:  true  .", " true ", true)]
    [InlineData("untrue / true", "true", true)]
    public async Task BooleanFieldsRequireACompleteObservedToken(string source, string raw, bool accepted) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("flag", OfficeAiFieldType.Boolean) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(accepted ? OfficeAiFieldStatus.Present : OfficeAiFieldStatus.Invalid, field.Status);
        Assert.Equal(accepted ? raw.Trim().ToLowerInvariant() : null, field.NormalizedValue);
        if (accepted) Assert.Equal(source.LastIndexOf(raw, StringComparison.Ordinal), Assert.Single(field.Citations).QuoteStart);
    }

    [Fact]
    public async Task BooleanQuoteCannotBorrowACompleteOccurrenceOutsideItsObservedFragment() {
        string source = "untrue " + new string('x', 100000) + " true";
        var executor = new Executor((sent, _) => {
            using var json = JsonDocument.Parse(sent.InputJson);
            var observation = json.RootElement.GetProperty("evidence")[0];
            string id = observation.GetProperty("id").GetString()!;
            string text = observation.GetProperty("text").GetString()!;
            string response = text.Contains("untrue", StringComparison.Ordinal) ? Field("true", id)
                : JsonSerializer.Serialize(new { claims = Array.Empty<object>(),
                    fields = new { field1 = new { status = "missing", rawValue = (string?)null, evidence = Array.Empty<object>() } },
                    blocks = Array.Empty<object>(), tables = Array.Empty<object>() });
            return Task.FromResult(new OfficeAiExecutionResponse(response));
        });
        var result = await new OfficeAiEngine(executor).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("flag", OfficeAiFieldType.Boolean) },
            Limits = new() { MaxRequestCharacters = 48000 }
        });
        Assert.True(executor.Requests.Count > 1);
        Assert.Equal(OfficeAiFieldStatus.Invalid, Assert.Single(result.Fields).Status);
    }
}
