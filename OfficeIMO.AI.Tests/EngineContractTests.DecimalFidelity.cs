using System.Text;
using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed partial class EngineContractTests {
    [Theory]
    [InlineData("0.123456789012345678901234567890", "en-US")]
    [InlineData("0.00000000000000000000000000001", "en-US")]
    [InlineData("-0.00000000000000000000000000001", "en-US")]
    [InlineData("0,123456789012345678901234567890", "de-DE")]
    public async Task DecimalExtractionRejectsLossOfExactSourceValue(string raw, string culture) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Equal(OfficeAiFieldStatus.Invalid, field.Status);
        Assert.Equal(raw, field.RawValue);
        Assert.Null(field.NormalizedValue);
        Assert.Equal(raw, Assert.Single(field.Citations).Quote);
    }

    [Theory]
    [InlineData("+001,234.5600", "en-US", "1234.5600")]
    [InlineData("1234.56-", "en-US", "-1234.56")]
    [InlineData("1\u202f234,5600", "fr-FR", "1234.5600")]
    [InlineData("0.100000000000000000000000000000", "en-US", "0.1000000000000000000000000000")]
    public async Task ExactDecimalsAllowGroupingAndInsignificantZeros(string raw, string culture, string expected) {
        var result = await new OfficeAiEngine(new Executor(Field(raw, "e1"))).RunAsync(Document(raw), Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = culture,
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) }
        });
        var field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Present, field.Status);
        Assert.Equal(expected, field.NormalizedValue);
    }

    [Fact]
    public async Task LossyDecimalsCannotHideConflictingSourceValuesAcrossBatches() {
        string[] values = { "0.00000000000000000000000000001", "0.00000000000000000000000000002" };
        var document = OfficeAiDocument.FromReadResult(Encoding.UTF8.GetBytes(string.Join("\n", values)), new OfficeDocumentReadResult {
            Blocks = values.Select((value, index) => new OfficeDocumentBlock {
                Text = value + " " + new string('x', 18000), Location = new() { Page = index + 1 }
            }).ToArray()
        });
        var executor = new Executor((request, _) => {
            using var input = JsonDocument.Parse(request.InputJson);
            var evidence = input.RootElement.GetProperty("evidence")[0];
            string raw = evidence.GetProperty("text").GetString()!.Split(' ')[0];
            return Task.FromResult(new OfficeAiExecutionResponse(Field(raw, evidence.GetProperty("id").GetString()!)));
        });
        var result = await new OfficeAiEngine(executor).RunAsync(document, Request() with {
            Operation = OfficeAiOperation.ExtractFields, Culture = "en-US",
            Fields = new[] { new OfficeAiFieldDefinition("amount", OfficeAiFieldType.Decimal) },
            Limits = new() { MaxRequestCharacters = 30000 }
        });
        Assert.Equal(2, executor.Requests.Count);
        var field = Assert.Single(result.Fields);
        Assert.Equal(OfficeAiFieldStatus.Conflicting, field.Status);
        Assert.Null(field.NormalizedValue);
        Assert.Null(field.RawValue);
        Assert.Equal(values, field.Citations.Select(citation => citation.Quote));
        Assert.Equal(OfficeAiResultStatus.Partial, result.Status);
        Assert.Contains("field-normalization-failed", result.Diagnostics);
    }

}
