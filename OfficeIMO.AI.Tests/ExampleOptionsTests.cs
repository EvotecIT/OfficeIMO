using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class ExampleOptionsTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ImageCapabilityIsIndependentOfRequestedImageInput(bool declared, bool requested) {
        var options = ExampleOptions.Parse(declared ? new[] { "--model-supports-images" } : Array.Empty<string>());
        if (requested && !declared) {
            Assert.Throws<ArgumentException>(() => ExampleExecution.BuildProfile(options, requested));
            return;
        }
        Assert.Equal(declared, ExampleExecution.BuildProfile(options, requested).SupportsImages);
    }

    [Fact]
    public void CopilotRequiresAnExplicitModelBeforeConnecting() {
        Assert.Throws<ArgumentException>(() => ExampleOptions.Parse(new[] { "--copilot" }));
        var options = ExampleOptions.Parse(new[] { "--copilot", "--model", "selected-model" });
        Assert.True(options.Copilot);
        Assert.Equal("selected-model", options.Model);
    }
}
