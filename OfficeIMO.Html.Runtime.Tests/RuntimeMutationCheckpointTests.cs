using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeMutationCheckpointTests {
    private static HtmlProcessRuntimeProvider Runtime()=>new(Path.Combine(AppContext.BaseDirectory,"RuntimeWorker","OfficeIMO.Html.Runtime.Worker.dll"),AngleSharpDomServices.Instance);

    [Fact]
    public async Task MutationNotificationsKeepTheirPositionAmongPromiseJobs() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.order=[];new MutationObserver(()=>order.push('observer')).observe(document.body,{attributes:true});Promise.resolve().then(()=>order.push('before'));document.body.setAttribute('data-test','yes');Promise.resolve().then(()=>order.push('after'))");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("before,observer,after",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task DocumentObservationIncludesDocumentChildrenAndRootReplacement() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.records=[];new MutationObserver(batch=>records.push(...batch.map(r=>({target:r.target===document,added:[...r.addedNodes].map(n=>n.nodeName),removed:[...r.removedNodes].map(n=>n.nodeName)})))).observe(document,{childList:true});document.appendChild(document.createComment('one'));const root=document.documentElement;root.remove();document.appendChild(root)");
        await session.WaitForAsync("records.length===3");
        Assert.True((await session.EvaluateAsync("records.every(r=>r.target) && records[0].added[0]==='#comment' && records[1].removed[0]==='HTML' && records[2].added[0]==='HTML'")).GetBoolean());
    }

    [Fact]
    public async Task PendingOrderFollowsFirstMutationAndOneCompoundNotification() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='a'></div><div id='b'></div>"});
        await session.ExecuteAsync("window.order=[];const a=document.querySelector('#a'),b=document.querySelector('#b');new MutationObserver(()=>order.push('a')).observe(a,{attributes:true});new MutationObserver(()=>{order.push('b');Promise.resolve().then(()=>order.push('job'))}).observe(b,{attributes:true});b.id='b2';a.id='a2'");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("b,a,job",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task LaterPendingObserversDrainRecordsAddedByEarlierCallbacks() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='a'></div><div id='b'></div>"});
        await session.ExecuteAsync("window.order=[];const a=document.querySelector('#a'),b=document.querySelector('#b');new MutationObserver(()=>{order.push('a');b.setAttribute('second','2');Promise.resolve().then(()=>order.push('job'))}).observe(a,{attributes:true});new MutationObserver(records=>order.push('b:'+records.length)).observe(b,{attributes:true});a.id='a2';b.id='b2'");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("a,b:2,job",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task ObserverCreatedDuringNotificationJoinsTheNextMicrotask() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='a'></div><div id='b'></div>"});
        await session.ExecuteAsync("window.order=[];const a=document.querySelector('#a'),b=document.querySelector('#b');new MutationObserver(()=>{order.push('a');Promise.resolve().then(()=>order.push('job'));new MutationObserver(()=>order.push('b')).observe(b,{attributes:true});b.id='b2'}).observe(a,{attributes:true});a.id='a2'");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("a,job,b",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Theory]
    [InlineData("child.remove()")]
    [InlineData("document.body.textContent=''")]
    public async Task DetachedSubtreeRemainsObservedUntilTheNotificationThenExpires(string removal) {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='child'><span>Text</span></div>"});
        await session.ExecuteAsync("window.names=[];window.child=document.querySelector('#child');new MutationObserver(records=>names.push(...records.map(r=>r.attributeName))).observe(document.body,{attributes:true,subtree:true});"+removal+";child.setAttribute('during','yes')");
        await session.WaitForAsync("names.length===1");
        await session.ExecuteAsync("child.setAttribute('after','ignored');window.finished=false;setTimeout(()=>finished=true,0)");
        await session.WaitForAsync("finished");
        Assert.Equal("during",(await session.EvaluateAsync("names.join(',')")).GetString());
    }

    [Fact]
    public async Task FragmentReplacementRecordsKeepAddedNodesAndBothSiblings() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='parent'><i id='before'></i><b id='old'></b><i id='after'></i></div>"});
        await session.ExecuteAsync("window.records=[];const parent=document.querySelector('#parent');new MutationObserver(batch=>records.push(...batch)).observe(parent,{childList:true});const fragment=document.createDocumentFragment();fragment.appendChild(document.createElement('u'));fragment.appendChild(document.createElement('em'));parent.replaceChild(fragment,document.querySelector('#old'))");
        await session.WaitForAsync("records.length===1");
        Assert.True((await session.EvaluateAsync("records[0].addedNodes.length===2 && records[0].addedNodes[0].nodeName==='U' && records[0].addedNodes[1].nodeName==='EM' && records[0].removedNodes[0].id==='old' && records[0].previousSibling.id==='before' && records[0].nextSibling.id==='after'")).GetBoolean());
    }

    [Fact]
    public async Task DocumentCanBeObservedWithoutAnElementAndAcrossReplacement() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("const root=document.documentElement;root.remove();window.records=[];new MutationObserver(batch=>records.push(...batch)).observe(document,{childList:true,subtree:true,attributes:true});document.appendChild(root);document.body.id='restored'");
        await session.WaitForAsync("records.length===2");
        Assert.True((await session.EvaluateAsync("records[0].target===document && records[1].target===document.body && records[1].attributeName==='id'")).GetBoolean());
    }

    [Fact]
    public async Task SecondaryDocumentMutationsShareTheSameAgentQueue() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.order=[];const other=document.implementation.createHTMLDocument('Secondary');new MutationObserver(()=>order.push('secondary')).observe(other,{subtree:true,attributes:true});Promise.resolve().then(()=>order.push('before'));other.body.id='changed';Promise.resolve().then(()=>order.push('after'))");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("before,secondary,after",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task DirectObservationSurvivesAdoptionIntoAnInertSecondaryDocument() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='target'></div>"});
        await session.ExecuteAsync("window.names=[];const node=document.querySelector('#target');new MutationObserver(records=>names.push(...records.map(r=>r.attributeName))).observe(node,{attributes:true});const other=document.implementation.createHTMLDocument('Secondary');other.body.appendChild(node);node.setAttribute('adopted','yes');const script=other.createElement('script');script.textContent='window.escaped=true';other.body.appendChild(script)");
        await session.WaitForAsync("names.length===1");
        Assert.True((await session.EvaluateAsync("names[0]==='adopted' && window.escaped===undefined")).GetBoolean());
    }

    [Fact]
    public async Task NativeAndExplicitMicrotasksDoNotDependOnMutablePromiseHelpers() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.order=[];new MutationObserver(()=>order.push('observer')).observe(document,{subtree:true,attributes:true});Object.defineProperty(Promise,Symbol.species,{get(){throw new Error('species override')}});Promise.prototype.then=()=>{throw new Error('then override')};Promise.resolve=()=>{throw new Error('resolve override')};document.body.id='changed';queueMicrotask(()=>order.push('explicit'))");
        await session.WaitForAsync("order.length===2");
        Assert.Equal("observer,explicit",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task InternalBaseTrackingDoesNotQueueAnEarlierWebNotification() {
        await using var session=await Runtime().OpenTrustedAsync(new());
        await session.ExecuteAsync("window.order=[];new MutationObserver(()=>order.push('observer')).observe(document.body,{attributes:true});document.head.appendChild(document.createElement('meta'));Promise.resolve().then(()=>order.push('before'));document.body.id='changed';Promise.resolve().then(()=>order.push('after'))");
        await session.WaitForAsync("order.length===3");
        Assert.Equal("before,observer,after",(await session.EvaluateAsync("order.join(',')")).GetString());
    }

    [Fact]
    public async Task MutationTargetsReportTheirCurrentTreeAfterRemoval() {
        await using var session=await Runtime().OpenTrustedAsync(new(){Html="<div id='child'></div>"});
        await session.ExecuteAsync("window.positions=[];const child=document.querySelector('#child');new MutationObserver(records=>positions.push(records[0].target.compareDocumentPosition(document.body))).observe(child,{attributes:true});child.id='changed';child.remove();window.documentOrder=document.compareDocumentPosition(document.body)");
        await session.WaitForAsync("positions.length===1");
        Assert.True((await session.EvaluateAsync("(positions[0]&Node.DOCUMENT_POSITION_DISCONNECTED)!==0 && (documentOrder&Node.DOCUMENT_POSITION_DISCONNECTED)===0 && (documentOrder&Node.DOCUMENT_POSITION_CONTAINED_BY)!==0")).GetBoolean(),(await session.EvaluateAsync("({positions,documentOrder,disconnected:Node.DOCUMENT_POSITION_DISCONNECTED,contained:Node.DOCUMENT_POSITION_CONTAINED_BY})")).ToString());
    }
}
