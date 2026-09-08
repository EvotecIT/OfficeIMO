using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class ExampleOptionsTests {
    [Fact]
    public void CopilotRequiresAnExplicitModelBeforeConnecting() {
        Assert.Throws<ArgumentException>(() => ExampleOptions.Parse(new[] { "--copilot" }));
        var options = ExampleOptions.Parse(new[] { "--copilot", "--model", "selected-model" });
        Assert.True(options.Copilot);
        Assert.Equal("selected-model", options.Model);
    }
}
