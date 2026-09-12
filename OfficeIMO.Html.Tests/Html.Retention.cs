using System;
using System.Threading;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlRetentionTests {
    private static HtmlDocument Parse(string source) => AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AttachedClonePreservesLiveIdentityAndDropsDetachedHistory(bool freeze) {
        HtmlDocument original = Parse("<!doctype html><p>Kept</p><aside>Removed</aside><template><b>Template</b></template>").Clone();
        HtmlElement removed = original.QuerySelector("aside")!;
        removed.Remove();
        HtmlElement kept = original.QuerySelector("p")!;
        HtmlElement template = original.QuerySelector("template")!;
        HtmlNode fragment = template.TemplateContent!;
        HtmlNode detached = original.CreateTextNode("Never attached");
        long revision = original.Revision;
        if (freeze) original.Freeze();

        HtmlDocument compact = original.CloneAttached();
        Assert.False(compact.IsReadOnly);
        Assert.NotEqual(original.SnapshotId, compact.SnapshotId);
        Assert.Equal(original.Mode, compact.Mode);
        Assert.Equal(original.ProviderId, compact.ProviderId);
        Assert.Equal(original.OuterHtml, compact.OuterHtml);
        Assert.Null(compact.GetNode(removed.NodeId));
        Assert.Null(compact.GetNode(removed.FirstChild!.NodeId));
        Assert.Null(compact.GetNode(detached.NodeId));
        Assert.Same(fragment, original.GetNode(fragment.NodeId));
        Assert.Equal("Template", compact.GetNode(fragment.NodeId)!.TextContent);
        Assert.Same(compact.GetNode(fragment.NodeId), ((HtmlElement)compact.GetNode(template.NodeId)!).TemplateContent);
        Assert.Equal(kept.SourceIndex, compact.GetNode(kept.NodeId)!.SourceIndex);
        compact.GetNode(kept.NodeId)!.TextContent = "Edited";
        Assert.True(compact.CreateElement("footer").NodeId > detached.NodeId);
        Assert.Equal("Kept", kept.TextContent);
        Assert.Equal(revision, original.Revision);
        Assert.NotNull(original.Clone().GetNode(removed.NodeId));
    }

    [Fact]
    public void CancelledAttachedCloneDoesNotChangeItsSource() {
        HtmlDocument original = Parse("<p>Kept</p>").Clone();
        HtmlNode detached = original.CreateTextNode("Detached");
        long revision = original.Revision;
        Assert.Throws<OperationCanceledException>(() => original.CloneAttached(new CancellationToken(true)));
        Assert.Equal(revision, original.Revision);
        Assert.Same(detached, original.GetNode(detached.NodeId));
        Assert.False(original.IsReadOnly);
        Assert.Equal("Kept", original.Body!.TextContent);
    }
}
