using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeMutationRegistrationTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task NestedDetachmentsPreserveInheritedObservationUntilNotification() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<div id='outer'><div id='inner'><span id='leaf'></span></div></div>" });
        await session.ExecuteAsync("window.names=[];const outer=document.querySelector('#outer'),inner=document.querySelector('#inner');window.leaf=document.querySelector('#leaf');new MutationObserver(records=>names.push(...records.map(r=>r.attributeName))).observe(document.body,{subtree:true,attributes:true});outer.remove();inner.remove();leaf.remove();leaf.setAttribute('during','yes')");
        Assert.Equal("during", (await session.EvaluateAsync("names.join(',')")).GetString());
        await session.ExecuteAsync("leaf.setAttribute('after','ignored')");
        Assert.Equal("during", (await session.EvaluateAsync("names.join(',')")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OverlappingDirectAndTransientOptionsPreserveAllRecordsAndOldValues(bool reobserveSource) {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<div id='inner' inherited='old' shared='previous'></div>" });
        await session.ExecuteAsync("window.records=[];const inner=document.querySelector('#inner');const observer=new MutationObserver(batch=>records.push(...batch.map(r=>({name:r.attributeName,old:r.oldValue}))));observer.observe(inner,{attributes:true,attributeFilter:['direct','shared']});observer.observe(document.body,{attributes:true,subtree:true,attributeOldValue:true,attributeFilter:['inherited','shared']});inner.remove();" +
            (reobserveSource ? "observer.observe(document.body,{attributes:true,subtree:true});" : "") +
            "inner.setAttribute('inherited','new');inner.setAttribute('shared','current');inner.setAttribute('direct','value')");
        Assert.Equal(reobserveSource ? "shared,direct" : "inherited,shared,direct", (await session.EvaluateAsync("records.map(r=>r.name).join(',')")).GetString());
        Assert.True((await session.EvaluateAsync(reobserveSource ? "records.every(r=>r.old===null)" : "records[0].old==='old' && records[1].old==='previous' && records[2].old===null")).GetBoolean());
    }

    [Fact]
    public async Task PendingObserversFollowStandardTargetThenAncestorTraversal() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<div id='target'></div>" });
        await session.ExecuteAsync("window.order=[];const target=document.querySelector('#target');new MutationObserver(()=>order.push('ancestor')).observe(document.body,{attributes:true,subtree:true});new MutationObserver(()=>order.push('target')).observe(target,{attributes:true});target.id='changed'");
        Assert.Equal("target,ancestor", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ReobservingRetainsPerNodeOrderWhileDisconnectingResetsIt() {
        await using var session = await Runtime().OpenTrustedAsync(new() { Html = "<div id='target'></div>" });
        await session.ExecuteAsync("window.order=[];const target=document.querySelector('#target');const first=new MutationObserver(()=>order.push('first'));first.observe(document.body,{attributes:true});const second=new MutationObserver(()=>order.push('second'));second.observe(target,{attributes:true});first.observe(target,{attributes:true});second.observe(target,{attributes:true,attributeOldValue:true});target.id='changed'");
        Assert.Equal("second,first", (await session.EvaluateAsync("order.join(',')")).GetString());
        await session.ExecuteAsync("order=[];second.disconnect();second.observe(target,{attributes:true});target.id='again'");
        Assert.Equal("first,second", (await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task SecondaryDocumentCloneRetainsAgentAndInertness() {
        await using var session = await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.order=[];const clone=document.implementation.createHTMLDocument('Secondary').cloneNode(true);new MutationObserver(()=>order.push('observer')).observe(clone,{subtree:true,attributes:true});Promise.resolve().then(()=>order.push('before'));clone.body.id='changed';Promise.resolve().then(()=>order.push('after'));const script=clone.createElement('script');script.textContent='window.escaped=true';clone.body.appendChild(script)");
        Assert.Equal("before,observer,after", (await session.EvaluateAsync("order.join(',')")).GetString());
        Assert.True((await session.EvaluateAsync("window.escaped===undefined")).GetBoolean());
    }
}
