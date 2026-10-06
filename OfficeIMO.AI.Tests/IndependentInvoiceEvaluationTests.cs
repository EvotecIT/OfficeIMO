using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class IndependentInvoiceEvaluationTests {
    [Fact]
    public async Task IndependentInvoiceValuesSurviveRealTextCaptureAndScalarValidation() {
        foreach (var item in EvaluationIndependentInvoiceCorpus.Create()) {
            var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
            using var input = new MemoryStream(item.Source);
            var document = await OfficeAiDocument.ReadAsync(reader, input, "invoice.txt");
            var result = await new OfficeAiEngine(new InvoiceXmlExecutor(item.Source)).RunAsync(document, item.Request);
            Assert.True(item.Gold.Score(result).ContractPassed, JsonSerializer.Serialize(new { result.Status, result.Diagnostics, result.Fields }));
            Assert.All(result.Fields, field => Assert.True(field.TextValueMatched));
        }
    }

    // Deterministic independent XML selection is a provider-boundary fixture, not model-quality proof.
    private sealed class InvoiceXmlExecutor(byte[] source) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "invoice-oracle", Provider = "fixture", Model = "xml", IsLocal = true };
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            using var input = JsonDocument.Parse(request.InputJson);
            var observations = input.RootElement.GetProperty("evidence").EnumerateArray().ToArray();
            var xml = XDocument.Parse(Encoding.UTF8.GetString(source));
            XNamespace c = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
            XNamespace a = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
            var root = xml.Root!;
            string[] values = [root.Element(c + "ID")!.Value, root.Element(c + "IssueDate")!.Value,
                root.Element(c + "DocumentCurrencyCode")!.Value, root.Element(a + "LegalMonetaryTotal")!.Element(c + "PayableAmount")!.Value,
                root.Element(a + "LegalMonetaryTotal")!.Element(c + "TaxExclusiveAmount")!.Value];
            var fields = values.Select((value, index) => new { key = "field" + (index + 1), value }).ToDictionary(item => item.key,
                item => new { status = "present", rawValue = item.value, evidence = new[] { new { id = observations.First(record => record.GetProperty("text").GetString()!.Contains(item.value, StringComparison.Ordinal)).GetProperty("id").GetString(), quote = item.value } } });
            return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>(), fields
            })));
        }
    }
}
