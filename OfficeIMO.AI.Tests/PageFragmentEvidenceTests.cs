using System.Text.Json;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class PageFragmentEvidenceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageSelectionUsesItsFragmentTextAndGeometry(bool roundTrip) {
        var source = new OfficeDocumentReadResult { Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Alpha. Beta." } },
            Pages = new[] {
                new OfficeDocumentPage { Number = 1, Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Alpha.", Region = new() { X = 1 } } } },
                new OfficeDocumentPage { Number = 2, Blocks = new[] { new OfficeDocumentBlock { Id = "source", Text = "Beta.", Region = new() { X = 2 } } } }
            } };
        if (roundTrip) source = OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(source));
        var document = OfficeAiDocument.FromReadResult(new byte[] { 1 }, source);
        Assert.Equal(new[] { "Alpha.", "Beta." }, document.Evidence.Select(item => item.Text));
        Assert.Equal(new double[] { 1, 2 }, document.Evidence.Select(item => item.Region!.X));
        var executor = new Executor();
        var result = await new OfficeAiEngine(executor).RunAsync(document, new() { Instruction = "Read page two", Pages = new[] { 2 } });
        Assert.Equal(OfficeAiResultStatus.Completed, result.Status);
        Assert.Equal("Beta.", Assert.Single(result.Claims).Text);
        Assert.Equal(2, Assert.Single(Assert.Single(result.Claims).Citations).Page);
    }

    private sealed class Executor : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile { get; } = new() { Id = "pages", Provider = "fixture", Model = "fixture", IsLocal = true };
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            using var input = JsonDocument.Parse(request.InputJson);
            var evidence = Assert.Single(input.RootElement.GetProperty("evidence").EnumerateArray());
            Assert.Equal("Beta.", evidence.GetProperty("text").GetString());
            return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = new[] { new { text = "Beta.", evidence = new[] { new { id = evidence.GetProperty("id").GetString(), quote = "Beta." } } } },
                fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
            })));
        }
    }
}
