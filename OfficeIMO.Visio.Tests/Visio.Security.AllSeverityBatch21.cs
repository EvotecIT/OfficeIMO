using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests;

public class VisioAllSeverityBatch21Tests {
    [Fact]
    public void NumericNetworkTitleIdIsReservedBeforeGeneratedConnectors() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        VisioDocument document = VisioDocument.Create(path)
            .NetworkDiagram("Network", network => network
                .Title("Network", id: "1")
                .Server("source", "Source", 0, 0)
                .Server("target", "Target", 1, 0)
                .Ethernet("source", "target"));

        VisioPage page = Assert.Single(document.Pages);
        Assert.Contains(page.Shapes, shape => shape.Id == "1");
        Assert.DoesNotContain(page.Connectors, connector => connector.Id == "1");

        document.Save();
        Assert.Empty(VisioValidator.Validate(path));
    }

    [Fact]
    public void NumericArchitectureCalloutIdIsReservedBeforeGeneratedConnectors() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
        VisioDocument document = VisioDocument.Create(path)
            .ArchitectureDiagram("Architecture", diagram => diagram
                .Service("source", "Source", 0, 0)
                .Service("target", "Target", 1, 0)
                .DataFlow("source", "target")
                .Callout("source", "1", "Important", VisioSide.Top));

        VisioPage page = Assert.Single(document.Pages);
        Assert.Contains(page.Shapes, shape => shape.Id == "1");
        Assert.DoesNotContain(page.Connectors, connector => connector.Id == "1");

        document.Save();
        Assert.Empty(VisioValidator.Validate(path));
    }
}
