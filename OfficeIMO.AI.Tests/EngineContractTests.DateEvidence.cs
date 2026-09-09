using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData("12030-04-031", "2030-04-03", "yyyy-MM-dd", "en-US", false)]
    [InlineData("é2030-04-03", "2030-04-03", "yyyy-MM-dd", "en-US", false)]
    [InlineData("2030-04-03\u0301", "2030-04-03", "yyyy-MM-dd", "en-US", false)]
    [InlineData("2030-04-03", "2030-04", "yyyy-MM", "en-US", false)]
    [InlineData("03.04.2030", "04.2030", "MM.yyyy", "pl-PL", false)]
    [InlineData("Date: 2030-04-03.", "2030-04-03", "yyyy-MM-dd", "en-US", true)]
    [InlineData("Date: 03.04.2030.", "03.04.2030", "dd.MM.yyyy", "pl-PL", true)]
    [InlineData("12030-04-031 / 2030-04-03", "2030-04-03", "yyyy-MM-dd", "en-US", true)]
    [InlineData("1|2030|04|03", "2030|04|03", "yyyy|MM|dd", "en-US", false)]
    [InlineData("2030|04|03|1", "2030|04|03", "yyyy|MM|dd", "en-US", false)]
    [InlineData("1::2030::04::03", "2030::04::03", "yyyy'::'MM'::'dd", "en-US", false)]
    [InlineData("2030::04::03::1", "2030::04::03", "yyyy\"::\"MM\"::\"dd", "en-US", false)]
    [InlineData("1|2030|04|03", "2030|04|03", @"yyyy\|MM\|dd", "en-US", false)]
    [InlineData("2030|04|03|1", "2030|04|03", "yyyy'|'MM'|'%d", "en-US", false)]
    [InlineData("Date: 2030|04|03.", "2030|04|03", "yyyy|MM|dd", "en-US", true)]
    [InlineData("Date: 2030::04::03.", "2030::04::03", "yyyy'::'MM'::'dd", "en-US", true)]
    [InlineData("Date: 2030|04|03.", "2030|04|03", "yyyy'|'MM'|'%d", "en-US", true)]
    [InlineData("Date: 2030|04|03.", "2030|04|03", @"yyyy\|MM\|dd", "en-US", true)]
    [InlineData("1※2030※04※03", "2030※04※03", "yyyy'※'MM'※'dd", "en-US", false)]
    [InlineData("Date: 2030※04※03.", "2030※04※03", "yyyy'※'MM'※'dd", "en-US", true)]
    [InlineData("1|2030|04|03 / 2030|04|03", "2030|04|03", "yyyy|MM|dd", "en-US", true)]
    public async Task DateFieldsRequireACompleteObservedValue(string source, string raw, string format, string culture, bool accepted) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, format) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(accepted ? OfficeAiFieldStatus.Present : OfficeAiFieldStatus.Invalid, field.Status);
        Assert.Equal(accepted ? "2030-04-03" : null, field.NormalizedValue);
        if (accepted) Assert.Equal(source.LastIndexOf(raw, StringComparison.Ordinal), Assert.Single(field.Citations).QuoteStart);
    }

    [Fact]
    public async Task DateQuoteCannotBorrowACompleteOccurrenceOutsideItsObservedFragment() {
        string source = "12030-04-031 " + new string('x', 100000) + " 2030-04-03";
        var executor = new Executor((sent, _) => {
            using var json = JsonDocument.Parse(sent.InputJson);
            var observation = json.RootElement.GetProperty("evidence")[0];
            string id = observation.GetProperty("id").GetString()!;
            string text = observation.GetProperty("text").GetString()!;
            string response = text.Contains("12030-04-031", StringComparison.Ordinal) ? Field("2030-04-03", id)
                : JsonSerializer.Serialize(new { claims = Array.Empty<object>(),
                    fields = new { field1 = new { status = "missing", rawValue = (string?)null, evidence = Array.Empty<object>() } },
                    blocks = Array.Empty<object>(), tables = Array.Empty<object>() });
            return Task.FromResult(new OfficeAiExecutionResponse(response));
        });
        var result = await new OfficeAiEngine(executor).RunAsync(Document(source), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Fields = new[] { new OfficeAiFieldDefinition("date", OfficeAiFieldType.Date, "yyyy-MM-dd") },
            Limits = new() { MaxRequestCharacters = 48000 }
        });
        Assert.True(executor.Requests.Count > 1);
        Assert.Equal(OfficeAiFieldStatus.Invalid, Assert.Single(result.Fields).Status);
    }
}
