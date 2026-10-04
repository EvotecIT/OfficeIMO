using System;
using System.IO;
using OfficeIMO.Data;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointDataProjectionContractsTests {
    [Theory]
    [InlineData(null, "Metrics.Score")]
    [InlineData("Custom", "Custom.Score")]
    [InlineData("", "Score")]
    public void ObjectTableUsesConfiguredCollectionHeaderPrefix(string? prefix, string expected) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        try {
            using var presentation = PowerPointPresentation.Create(path);
            var table = presentation.AddSlide().AddTable(new[] { new Row() }, options => {
                options.Columns = new[] { "Metrics.Score" };
                options.CollectionMapColumns["Metrics"] = new CollectionColumnMapping { HeaderPrefix = prefix };
            });
            Assert.Equal(expected, table.GetCell(0, 0).Text);
            Assert.Equal("12", table.GetCell(1, 0).Text);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
    private sealed class Row { public Metric[] Metrics { get; } = new[] { new Metric() }; }
    private sealed class Metric { public string Name => "Score"; public int Value => 12; }
}
