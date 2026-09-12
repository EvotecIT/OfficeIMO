using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeHistoryTests {
    private static readonly Uri Page = new("https://routing.example/app/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory,"RuntimeWorker","OfficeIMO.Html.Runtime.Worker.dll"),AngleSharpDomServices.Instance);

    [Fact]
    public async Task PushReplaceAndTraversalKeepStateSnapshotsAndRouteIdentity() {
        await using var session = await Runtime().OpenTrustedAsync(new() {DocumentUrl=Page});
        await session.ExecuteAsync("""
            window.events=[];window.original={page:'first'};
            history.replaceState(original,'','/first');original.page='changed';
            window.firstState=history.state;
            history.pushState({page:'second'},'','/second?q=2');
            onpopstate=e=>{events.push(e.state.page);window.same=e.state===history.state};
            history.back();window.synchronous=location.pathname;
            """);
        await session.WaitForAsync("events.length===1");
        Assert.True((await session.EvaluateAsync("synchronous==='/second' && location.pathname==='/first' && document.URL===location.href && firstState.page==='first' && same && history.state!==firstState && history.length===2")).GetBoolean());
        var first=await session.CaptureAsync();
        await session.ExecuteAsync("history.forward()");
        await session.WaitForAsync("events.length===2");
        Assert.Equal(new Uri("https://routing.example/first"),first.DocumentUrl);
        Assert.Equal(new Uri("https://routing.example/second?q=2"),(await session.CaptureAsync()).DocumentUrl);
    }

    [Fact]
    public async Task StateClonesCyclesCollectionsBinaryAliasesAndNativeValues() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page});
        await session.ExecuteAsync("""
            const buffer=new ArrayBuffer(8),view=new Uint16Array(buffer,2,2),data=new DataView(buffer,2,4);view[0]=42;
            const key={id:1},state={buffer,view,data,key,map:new Map([[key,42]]),set:new Set([key]),date:new Date(1234),regex:/ab/gi,big:12345678901234567890n,number:NaN,empty:undefined};
            state.self=state;state.arr=[,key];state.arr.extra='value';state.error=new TypeError('bad',{cause:key});
            history.replaceState(state,'');view[0]=7;key.id=99;
            window.result=history.state;
            """);
        Assert.True((await session.EvaluateAsync("result.self===result && result.key.id===1 && result.map.get(result.key)===42 && result.set.has(result.key) && result.view.buffer===result.buffer && result.data.buffer===result.buffer && result.view.byteOffset===2 && result.view[0]===42 && result.date.getTime()===1234 && result.regex.source==='ab' && result.regex.flags==='gi' && result.big===12345678901234567890n && Number.isNaN(result.number) && 'empty' in result && !(0 in result.arr) && result.arr[1]===result.key && result.arr.extra==='value' && result.error instanceof TypeError && result.error.cause===result.key")).GetBoolean());
    }

    [Theory]
    [InlineData("()=>1")]
    [InlineData("Symbol('x')")]
    [InlineData("new Proxy({},{})")]
    [InlineData("document.body")]
    [InlineData("Promise.resolve(1)")]
    public async Task UncloneableStateDoesNotChangeHistory(string value) {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page});
        await session.ExecuteAsync("window.failure='';try{history.pushState("+value+",'', '/bad')}catch(e){failure=e.name}");
        Assert.Equal("DataCloneError",(await session.EvaluateAsync("failure")).GetString());
        Assert.True((await session.EvaluateAsync("history.length===1 && location.pathname==='/app/start' && history.state===null")).GetBoolean());
    }

    [Fact]
    public async Task FragmentAssignmentQueuesEventsAndSupportsBackForwardAndReplacement() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page});
        await session.ExecuteAsync("window.events=[];onhashchange=e=>events.push({type:e.type,old:e.oldURL,next:e.newURL});onpopstate=e=>events.push({type:e.type,state:e.state});location.hash='details';window.immediate=events.length");
        await session.WaitForAsync("events.length===2");
        Assert.Equal(1,(await session.EvaluateAsync("immediate")).GetInt32());
        Assert.True((await session.EvaluateAsync("events[1].next===location.href && history.length===2 && location===document.location")).GetBoolean());
        await session.ExecuteAsync("location.replace('#other')");
        await session.WaitForAsync("events.length===4");
        await session.ExecuteAsync("history.back()");
        await session.WaitForAsync("events.length===6");
        Assert.True((await session.EvaluateAsync("events[4].type==='popstate' && events[5].type==='hashchange' && location.hash==='' && history.length===2")).GetBoolean(),(await session.EvaluateAsync("({events,hash:location.hash,length:history.length,url:location.href})")).ToString());
    }

    [Fact]
    public async Task RejectedNavigationAndStateBudgetsAreAtomic() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,MaxHistoryEntries=3,MaxHistoryStateBytes=200,MaxHistoryTotalStateBytes=200});
        await session.ExecuteAsync("window.failures=[];history.replaceState('a'.repeat(70),'');try{history.pushState('b'.repeat(70),'','/too-big')}catch(e){failures.push(e.name)};try{history.pushState(null,'','https://other.example/')}catch(e){failures.push(e.name)};try{location.href='/new-document'}catch(e){failures.push(e.name)}");
        Assert.True((await session.EvaluateAsync("failures.join(',')==='QuotaExceededError,SecurityError,NotSupportedError' && history.length===1 && history.state.length===70 && location.pathname==='/app/start'")).GetBoolean());
        await session.ExecuteAsync("history.pushState(null,'','/two');history.pushState(null,'','/three');history.pushState(null,'','/four');window.popped=false;onpopstate=()=>popped=true;history.go(-2)");
        await session.WaitForAsync("popped");
        Assert.True((await session.EvaluateAsync("history.length===3 && location.pathname==='/app/start' && history.state.length===70")).GetBoolean());
    }

    [Fact]
    public async Task RouteChangesUpdateRelativeFetchAndImportBasesButRespectBaseElements() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Resources=new[]{
            HtmlRuntimeResource.FromText(new Uri(Page,"/moved/data.json"),"{\"value\":42}","application/json"),
            HtmlRuntimeResource.FromText(new Uri(Page,"/assets/module.js"),"export const value=7","text/javascript")
        }});
        await session.ExecuteAsync("history.pushState(null,'','/moved/page');fetch('data.json').then(r=>r.json()).then(r=>window.result=r.value)");
        await session.WaitForAsync("window.result===42");
        await session.ExecuteAsync("document.head.innerHTML='<base href=\"/assets/\">'");
        await session.ExecuteAsync("import('./module.js').then(m=>window.module=m.value)");
        await session.WaitForAsync("window.module===7");
        Assert.Equal("https://routing.example/assets/",(await session.EvaluateAsync("document.baseURI")).GetString());
    }

    [Fact]
    public async Task LinksAndScriptClicksShareCancellationAndFragmentActivation() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<a id='route' href='#details'><span>Details</span></a><input type='checkbox'><button hidden>Hidden</button>"});
        await session.ExecuteAsync("window.events=[];onhashchange=e=>events.push(e instanceof HashChangeEvent && e.isTrusted);document.querySelector('#route').onclick=e=>e.preventDefault()");
        await session.Locator("#route span").ClickAsync();
        Assert.Equal("",(await session.EvaluateAsync("location.hash")).GetString());
        await session.ExecuteAsync("document.querySelector('#route').onclick=null;document.querySelector('#route').click()");
        await session.WaitForAsync("events.length===1");
        Assert.True((await session.EvaluateAsync("events[0] && location.hash==='#details' && history instanceof History && location instanceof Location")).GetBoolean());
        await session.ExecuteAsync("window.clicks=0;const button=document.querySelector('button');button.onclick=()=>{clicks++;button.click()};button.click();document.querySelector('input').click()");
        Assert.True((await session.EvaluateAsync("clicks===1 && document.querySelector('input').checked")).GetBoolean());
    }

    [Fact]
    public async Task PendingTaskLimitRefusesFragmentChangesBeforeUpdatingTheRoute() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,MaxPendingHistoryTasks=1});
        await session.ExecuteAsync("window.failure='';history.back();try{location.hash='not-committed'}catch(e){failure=e.name};window.unchanged=location.hash===''");
        Assert.True((await session.EvaluateAsync("failure==='QuotaExceededError' && unchanged && history.length===1")).GetBoolean());
    }

    [Fact]
    public async Task BaseElementsFreezeTheirUrlAcrossHistoryAndReleaseItWhenRemoved() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<base href='./assets/'><a href='item'>Item</a>"});
        await session.ExecuteAsync("history.pushState(null,'','/other/page');window.frozen=document.baseURI;document.querySelector('base').remove();window.updated=document.querySelector('a').href");
        Assert.True((await session.EvaluateAsync("frozen==='https://routing.example/app/assets/' && updated==='https://routing.example/other/item'")).GetBoolean());
    }

    [Fact]
    public async Task ModuleRouterUsesHistoryFetchAndActionsThenConvertsIndependentCaptures() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,
            Html="<h1>Routed report</h1><nav><a href='/reports/details/'>Details</a></nav><p id='view'>Loading</p><script type='module' src='/router.js'></script>",
            Resources=new[]{
                HtmlRuntimeResource.FromText(new Uri(Page,"/router.js"),"""
                    history.replaceState({page:'overview'},'');
                    async function render(){
                        if(history.state.page==='overview'){document.querySelector('#view').textContent='Overview';return;}
                        const data=await fetch('data.json').then(r=>r.json());
                        document.querySelector('#view').textContent='Details total: '+data.total;
                    }
                    document.querySelector('a').onclick=e=>{e.preventDefault();history.pushState({page:'details'},'','/reports/details/');render()};
                    onpopstate=e=>{if(!(e instanceof PopStateEvent)||!e.isTrusted)throw new Error('Invalid traversal event');render()};
                    await render();
                    ""","text/javascript"),
                HtmlRuntimeResource.FromText(new Uri(Page,"/reports/details/data.json"),"{\"total\":42}","application/json")
            }});
        await session.Locator("a").ClickAsync();
        await session.Locator("#view").WaitForTextAsync("Details total: 42");
        var details=await session.CaptureAsync();
        await session.ExecuteAsync("history.back()");
        await session.Locator("#view").WaitForTextAsync("Overview");
        var overview=await session.CaptureAsync();
        await session.DisposeAsync();
        Assert.Equal(new Uri(Page,"/reports/details/"),details.DocumentUrl);
        Assert.Equal(Page,overview.DocumentUrl);
        var conversion=HtmlConversionDocument.FromDocument(details.Document,new(){BaseUri=details.DocumentUrl});
        Assert.Contains("Details total: 42",conversion.ToMarkdown());
        Assert.Contains("Details total: 42",PdfReadDocument.Open(conversion.ToPdfBytes()).ExtractText());
    }

    [Fact]
    public async Task RedispatchedHistoryEventsLoseTrustedStatusAndRetainTheirPayload() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page});
        await session.ExecuteAsync("window.trust=[];onhashchange=e=>{try{dispatchEvent(e)}catch(_){}trust.push(e.isTrusted);window.last=e};location.hash='details'");
        await session.WaitForAsync("trust.length===1");
        await session.ExecuteAsync("dispatchEvent(last)");
        Assert.True((await session.EvaluateAsync("trust[0]===true && trust[1]===false && last instanceof HashChangeEvent && last.newURL===location.href")).GetBoolean());
    }

    [Fact]
    public async Task HistoryBudgetsAndCloningSurviveChangesToPublicCollectionMethods() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,MaxHistoryStateBytes=200,MaxHistoryTotalStateBytes=200});
        await session.ExecuteAsync("history.replaceState('a'.repeat(70),'');window.failure='';Array.prototype.reduce=()=>0;Array.prototype.slice=()=>[];Map.prototype.has=()=>true;Map.prototype.get=()=>({bypass:true});try{History.prototype.pushState.call(history,'b'.repeat(70),'','/bypass')}catch(e){failure=e.name};window.intact=history.length===1 && location.pathname==='/app/start'");
        Assert.True((await session.EvaluateAsync("failure==='QuotaExceededError' && intact")).GetBoolean());
    }

    [Fact]
    public async Task ClearingFragmentsAndTraversingRunPromiseJobsBeforeHashChange() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=new Uri(Page.AbsoluteUri+"#one")});
        await session.ExecuteAsync("window.events=[];onpopstate=()=>{events.push('pop');Promise.resolve().then(()=>events.push('micro'))};onhashchange=()=>events.push('hash');location.hash='';events.push('sync')");
        await session.WaitForAsync("events.length===4");
        Assert.Equal("pop,sync,micro,hash",(await session.EvaluateAsync("events.join(',')")).GetString());
        Assert.Equal(Page.AbsoluteUri+"#",(await session.EvaluateAsync("location.href")).GetString());
        await session.ExecuteAsync("history.pushState({page:2},'','#two');events=[];history.back()");
        await session.WaitForAsync("events.length===3");
        Assert.Equal("pop,micro,hash",(await session.EvaluateAsync("events.join(',')")).GetString());
    }

    [Fact]
    public async Task ANewBaseImmediatelyAffectsDynamicallyInsertedScripts() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Resources=new[]{
            HtmlRuntimeResource.FromText(new Uri(Page,"/assets/inserted.js"),"window.inserted=42","text/javascript")
        }});
        await session.ExecuteAsync("document.head.innerHTML='<base href=\"/assets/\">';const script=document.createElement('script');script.src='inserted.js';document.body.appendChild(script)");
        await session.WaitForAsync("window.inserted===42");
        Assert.Equal("https://routing.example/assets/",(await session.EvaluateAsync("document.querySelector('base').href")).GetString());
    }

    [Fact]
    public async Task TraversalPreservesUndefinedStateAndFragmentEventsReserveTheirQueueSlot() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,MaxPendingHistoryTasks=1});
        await session.ExecuteAsync("history.replaceState(undefined,'');history.pushState(1,'','/two');window.popped=false;onpopstate=e=>{window.exact=e.state===undefined && history.state===undefined;popped=true};history.back()");
        await session.WaitForAsync("popped");
        Assert.True((await session.EvaluateAsync("exact")).GetBoolean());
        await session.ExecuteAsync("window.events=[];onpopstate=()=>{try{history.back()}catch(e){events.push(e.name)}};onhashchange=()=>events.push('hash');location.hash='one'");
        await session.WaitForAsync("events.length===2");
        Assert.True((await session.EvaluateAsync("events.join(',')==='QuotaExceededError,hash' && location.hash==='#one'")).GetBoolean());
    }

    [Theory]
    [InlineData("base.remove();document.head.appendChild(base)")]
    [InlineData("base.href='./temporary/';base.href='./assets/'")]
    [InlineData("base.href='./assets/'")]
    [InlineData("const earlier=document.createElement('base');earlier.href='/temporary/';base.before(earlier);earlier.remove()")]
    public async Task BaseMutationTransitionsRefreezeAgainstTheCurrentRoute(string mutation) {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<base href='./assets/'><a href='item'>Item</a>"});
        await session.ExecuteAsync("history.pushState(null,'','/other/page');const base=document.querySelector('base');"+mutation);
        Assert.Equal("https://routing.example/other/assets/item",(await session.EvaluateAsync("document.querySelector('a').href")).GetString());
    }

    [Fact]
    public async Task DetachedRadioActivationStaysWithinItsOwnTree() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<input type='radio' name='choice' checked>"});
        await session.ExecuteAsync("const holder=document.createElement('div');holder.innerHTML='<input type=radio name=choice checked><input type=radio name=choice>';window.peer=holder.firstChild;window.target=holder.lastChild;target.click();window.isolated=document.createElement('input');isolated.type='radio';isolated.name='choice';isolated.click()");
        Assert.True((await session.EvaluateAsync("document.querySelector('input').checked && !peer.checked && target.checked && isolated.checked")).GetBoolean());
    }

    [Fact]
    public async Task CapturesPreserveTheFrozenBaseForInertConversion() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<base href='./assets/'><a href='item'>Item</a>"});
        await session.ExecuteAsync("history.pushState(null,'','/other/page')");
        var capture=await session.CaptureAsync();
        Assert.Equal("./assets/",capture.Document.QuerySelector("base")!.GetAttribute("href"));
        Assert.Equal(new Uri(Page,"/app/assets/"),capture.BaseUri);
        var conversion=HtmlConversionDocument.FromDocument(capture.CreateStandaloneDocument(),new(){BaseUri=capture.DocumentUrl});
        Assert.Equal(new Uri(Page,"/app/assets/"),conversion.BaseUri);
        Assert.Contains("https://routing.example/app/assets/item",conversion.ToMarkdown());
    }

    [Fact]
    public async Task UnrelatedBaseAndBodyMutationsPreserveTheFrozenBase() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<base href='./assets/'><a href='item'>Item</a>"});
        await session.ExecuteAsync("history.pushState(null,'','/other/page');const later=document.createElement('base');later.href='/unused/';document.head.appendChild(later);later.remove();document.body.appendChild(document.createElement('p'))");
        Assert.Equal("https://routing.example/app/assets/",(await session.EvaluateAsync("document.baseURI")).GetString());
    }

    [Fact]
    public async Task StandaloneCapturesSupplyAnAbsoluteBaseWithoutChangingAuthoredMarkup() {
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html="<a href='item'>Item</a>"});
        var capture=await session.CaptureAsync();
        var standalone=capture.CreateStandaloneDocument();
        Assert.Null(capture.Document.QuerySelector("base"));
        Assert.True(standalone.IsReadOnly);
        Assert.Equal(Page.AbsoluteUri,standalone.QuerySelector("base")!.GetAttribute("href"));
        Assert.Contains("https://routing.example/app/item",HtmlConversionDocument.FromDocument(standalone).ToMarkdown());
    }

    [Theory]
    [InlineData(true,"https://routing.example/other/assets/item")]
    [InlineData(false,"https://routing.example/app/assets/item")]
    public async Task TemporaryHrefActivationOnlyRefreezesAnEarlierBase(bool earlier,string expected) {
        string inactive="<base id='inactive'>",active="<base href='./assets/'>";
        await using var session=await Runtime().OpenTrustedAsync(new(){DocumentUrl=Page,Html=(earlier?inactive+active:active+inactive)+"<a href='item'>Item</a>"});
        await session.ExecuteAsync("history.pushState(null,'','/other/page');const inactive=document.querySelector('#inactive');inactive.setAttribute('href','/temporary/');inactive.removeAttribute('href')");
        Assert.Equal(expected,(await session.EvaluateAsync("document.querySelector('a').href")).GetString());
    }
}
