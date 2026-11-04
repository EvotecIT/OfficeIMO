using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordSmartArtNodeTests {
    [Theory]
    [InlineData(WordSmartArtType.BasicProcess)]
    [InlineData(WordSmartArtType.Cycle)]
    public void AddedNodeRemainsEditableAfterReorderingAndReloading(WordSmartArtType layout) {
        using var document = WordDocument.Create();
        WordSmartArt diagram = document.AddSmartArt(layout);
        diagram.ClearNodes();
        WordSmartArtNode first = diagram.AddNode("First");
        WordSmartArtNode second = diagram.AddNode("Second");
        Assert.Same(diagram, first.SmartArt);
        Assert.NotEqual(first.Id, second.Id);
        diagram.MoveNode(0, 1);
        first.Text = "Edited first";
        Assert.Equal(new[] { "Second", "Edited first" }, diagram.Nodes.Select(node => node.Text));

        using var reloaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordSmartArt loaded = Assert.Single(reloaded.SmartArts);
        Assert.Equal(first.Id, loaded.Nodes[1].Id);
        loaded.Nodes[1].Text = "Edited after reload";
        Assert.Equal("Edited after reload", loaded.GetNodeText(1));
        loaded.RemoveNodeAt(0);
        Assert.Equal("Edited after reload", loaded.Nodes[0].Text);

        diagram.RemoveNodeAt(1);
        Assert.Throws<InvalidOperationException>(() => first.Text);
        Assert.Throws<InvalidOperationException>(() => first.Text = "Detached");
        Assert.Equal("Second", second.Text);
    }
}
