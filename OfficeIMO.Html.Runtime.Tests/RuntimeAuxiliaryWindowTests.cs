using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeAuxiliaryWindowTests {
    private static readonly Uri Start = new("https://popup.example/reports/start");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task PopupSurvivesReloadWithItsRealmAndCapturedOpener()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            MaxAuxiliaryWindows = 1, Html = "<!doctype html><body><h1>Root</h1>"
        });
        await session.ExecuteAsync("""
            window.popup=open('','persistent');
            const script=popup.document.createElement('script');
            script.textContent="const originalOpener=opener;let count=0;addEventListener('message',()=>{count++;originalOpener.document.body.setAttribute('data-popup-count',String(count));originalOpener.postMessage({same:originalOpener===opener},'*');});opener.document.body.setAttribute('data-popup-ready','yes')";
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-popup-ready')");
        await session.ReloadAsync();
        await session.ExecuteAsync("""
            window.popup=open('','persistent');window.reply=null;
            addEventListener('message',e=>reply={same:e.data.same,source:e.source===popup});
            popup.postMessage('next','*');
            """);
        await session.WaitForAsync("reply!==null");
        Assert.True((await session.EvaluateAsync("reply.same && reply.source && !popup.closed && popup.opener===window && document.body.getAttribute('data-popup-count')==='1'")).GetBoolean());
        await session.ExecuteAsync("popup.close()");
        Assert.True((await session.EvaluateAsync("popup.closed && open('','another')===null")).GetBoolean());
    }

    [Fact]
    public async Task CrossOriginRootNavigationRevokesCapturedOpenerDomAccessButKeepsMessaging()
    {
        var other = new Uri("https://other.example/next");
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            ResourcePolicy = new() { AllowedOrigins = [new Uri("https://other.example/")] },
            Resources = [HtmlRuntimeResource.FromText(other, "<!doctype html><body>Other", "text/html"),
                HtmlRuntimeResource.FromText(new Uri(Start, "/data"), "popup-origin", "text/plain")],
            Html = "<!doctype html><body>Original"
        });
        await session.ExecuteAsync("""
            window.popup=open('','persistent');
            const script=popup.document.createElement('script');
            script.textContent="const saved=opener;addEventListener('message',()=>{let blocked=false;try{saved.document.body.textContent='leak'}catch(e){blocked=e.name==='SecurityError'}fetch('/data',{mode:'same-origin'}).then(r=>r.text()).then(value=>saved.postMessage({blocked,same:saved===opener,value},'*'))});opener.document.body.setAttribute('data-ready','yes')";
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-ready')");
        await session.NavigateAsync(other);
        await session.ExecuteAsync("""
            window.popup=open('','persistent');window.reply=null;
            addEventListener('message',event=>reply={...event.data,origin:event.origin,source:event.source===popup});
            window.blocked=false;try{popup.document.body.textContent='leak'}catch(e){blocked=e.name==='SecurityError'}
            popup.postMessage('check','https://popup.example');
            """);
        await session.WaitForAsync("reply!==null");
        Assert.True((await session.EvaluateAsync("blocked && reply.blocked && reply.same && reply.source && reply.value==='popup-origin' && reply.origin==='https://popup.example' && document.body.textContent==='Other'")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RootReloadRetiresOnlyRootCallbacksAndPreservesPopupTimers(bool scheduleOnPopup)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body>Root"
        });
        await session.ExecuteAsync("window.scheduleOnPopup=" + (scheduleOnPopup ? "true" : "false"));
        await session.ExecuteAsync("""
            window.popup=open('','clock');popup.document.body.setAttribute('data-root','0');
            let rootTicks=0;(scheduleOnPopup ? popup : window).setInterval(()=>popup.document.body.setAttribute('data-root',String(++rootTicks)),1);
            const script=popup.document.createElement('script');
            script.textContent="let ticks=0;setInterval(()=>document.body.setAttribute('data-popup',String(++ticks)),1);opener.document.body.setAttribute('data-ready','yes')";
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-ready') && Number(popup.document.body.getAttribute('data-popup'))>2 && Number(popup.document.body.getAttribute('data-root'))>2");
        await session.ReloadAsync();
        await session.ExecuteAsync("window.popup=open('','clock');window.rootTicks=popup.document.body.getAttribute('data-root');window.popupTicks=Number(popup.document.body.getAttribute('data-popup'))");
        await session.WaitForAsync("Number(popup.document.body.getAttribute('data-popup'))>popupTicks+5");
        Assert.True((await session.EvaluateAsync("popup.document.body.getAttribute('data-root')===rootTicks")).GetBoolean());
    }

    [Fact]
    public void AuxiliaryWindowLimitIsValidatedAndSnapshotted()
    {
        var request = new HtmlScriptRequest { MaxAuxiliaryWindows = 3 };
        var snapshot = request.Snapshot();
        request.MaxAuxiliaryWindows = 0;
        Assert.Equal(3, snapshot.MaxAuxiliaryWindows);
        Assert.Equal(0, request.Snapshot().MaxAuxiliaryWindows);
        request.MaxAuxiliaryWindows = -1;
        Assert.Throws<ArgumentOutOfRangeException>(request.Snapshot);
        request.MaxAuxiliaryWindows = 129;
        Assert.Throws<ArgumentOutOfRangeException>(request.Snapshot);
    }

    [Fact]
    public async Task InitialPopupKeepsWindowIdentityAndTheInitiatorsFrozenBase()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><base href='/assets/'><h1>Opener</h1>"
        });
        Assert.True((await session.EvaluateAsync("""
            (()=>{
                window.popup=open('','report');
                const same=popup===popup.document.defaultView && popup.parent===popup && popup.top===popup && popup.opener===window;
                document.querySelector('base').href='/later/';
                return same && popup.document.URL==='about:blank' && popup.document.baseURI==='https://popup.example/assets/' &&
                    popup.document.readyState==='complete' && open('','report')===popup;
            })()
            """)).GetBoolean());
        Assert.True((await session.EvaluateAsync("popup.document.body!==null && !popup.closed")).GetBoolean());
    }

    [Fact]
    public async Task DocumentWindowOverloadPreservesTheOpenerAndSupportsMessages()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><h1>Opener</h1>"
        });
        await session.ExecuteAsync("""
            window.popup=document.open('about:blank','report','');window.reply=null;
            window.addEventListener('message',event=>{reply={data:event.data,origin:event.origin,source:event.source===popup};});
            const script=popup.document.createElement('script');
            script.textContent="addEventListener('message',event=>opener.postMessage({value:event.data.value+1},'https://popup.example'));opener.document.body.setAttribute('data-ready','yes');";
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-ready')");
        await session.ExecuteAsync("popup.postMessage({value:41},'https://popup.example')");
        await session.WaitForAsync("reply!==null");
        Assert.True((await session.EvaluateAsync("reply.data.value===42 && reply.origin==='https://popup.example' && reply.source && document.querySelector('h1').textContent==='Opener'")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NamedReuseUsesTheCurrentNameOfEveryAuxiliaryWindow(bool initiallyNamed)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            MaxAuxiliaryWindows = 1, Html = "<!doctype html><body>Parent"
        });
        await session.ExecuteAsync(initiallyNamed ? "window.popup=open('','old')" : "window.popup=open()");
        await session.ExecuteAsync("popup.name='renamed'");
        Assert.True((await session.EvaluateAsync("open('','renamed')===popup")).GetBoolean());
        Assert.True((await session.EvaluateAsync("open('','old')===null")).GetBoolean());
        await session.ExecuteAsync("popup.name=''");
        Assert.True((await session.EvaluateAsync("open('','renamed')===null")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InactiveDocumentCannotOpenAnAuxiliaryWindow(bool closedPopup)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            MaxAuxiliaryWindows = closedPopup ? 2 : 1, Html = "<!doctype html><body>Parent"
        });
        await session.ExecuteAsync(closedPopup
            ? "window.first=open();window.stale=first.document;first.close();"
            : "const frame=document.createElement('iframe');document.body.append(frame);window.stale=frame.contentDocument;frame.remove();");
        Assert.Equal("InvalidAccessError", (await session.EvaluateAsync("(()=>{try{stale.open('','rejected','');return 'accepted'}catch(error){return error.name}})()")).GetString());
        Assert.True((await session.EvaluateAsync("open()!==null")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SandboxedFrameOpeningUsesItsPermissionAndInitiator(bool allowPopups)
    {
        const string child = "<base href='/child-assets/'><script>const popup=open('','child-report');parent.document.body.setAttribute('data-popup',popup===null?'blocked':popup.opener===window && popup.parent===popup && popup.document.baseURI===document.baseURI?'inherited':'wrong')</script>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, MaxAuxiliaryWindows = 1,
            Html = "<!doctype html><body><iframe sandbox='allow-scripts allow-same-origin" + (allowPopups ? " allow-popups" : "") + "' srcdoc=\"" + System.Net.WebUtility.HtmlEncode(child) + "\"></iframe>"
        });
        await session.WaitForAsync("document.body.hasAttribute('data-popup')");
        Assert.Equal(allowPopups ? "inherited" : "blocked", (await session.EvaluateAsync("document.body.getAttribute('data-popup')")).GetString());
        Assert.Equal(allowPopups, (await session.EvaluateAsync("open()===null")).GetBoolean());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task FormSubmissionKeepsTheInitiatorsSandbox(bool popup, bool allowForms)
    {
        string child = "<body><script>const target=" + (popup ? "open('','sandboxed-form').document" : "document") +
            ";target.body.innerHTML='<form action=https://popup.example/submit><input name=value value=42></form>';parent.document.body.setAttribute('data-ready','yes')</script>";
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<body><iframe sandbox='allow-scripts allow-same-origin allow-popups" + (allowForms ? " allow-forms" : "") +
                "' srcdoc=\"" + System.Net.WebUtility.HtmlEncode(child) + "\"></iframe>"
        });
        await session.WaitForAsync("document.body.hasAttribute('data-ready')");
        await session.ExecuteAsync("window.targetForm=" + (popup ? "open('','sandboxed-form').document" : "document.querySelector('iframe').contentDocument") + ".querySelector('form')");
        Assert.Equal(allowForms ? 1 : 0, (await session.EvaluateAsync("""
            (()=>{let submitted=0;targetForm.addEventListener('submit',event=>{event.preventDefault();submitted++});
                HTMLFormElement.prototype.requestSubmit.call(targetForm);return submitted;})()
            """)).GetInt32());
        if (!allowForms) {
            Assert.True((await session.EvaluateAsync("""
                (()=>{
                    const doc=targetForm.ownerDocument,foreign=doc.createElement('button'),invalid=doc.createElement('div');
                    const other=doc.createElement('form');other.append(foreign);doc.body.append(other);
                    return [foreign,invalid].every(submitter=>{try{HTMLFormElement.prototype.requestSubmit.call(targetForm,submitter);return false}catch(error){return error.name==='TypeError'}});
                })()
                """)).GetBoolean());
            Assert.True((await session.EvaluateAsync("""
                (()=>{
                    let invalid=0;const input=targetForm.querySelector('input');input.required=true;input.value='';
                    input.addEventListener('invalid',()=>invalid++);
                    const url=targetForm.ownerDocument.URL,rootUrl=document.URL;
                    HTMLFormElement.prototype.requestSubmit.call(targetForm);
                    HTMLFormElement.prototype.submit.call(targetForm);
                    return invalid===0 && targetForm.ownerDocument.URL===url && document.URL===rootUrl;
                })()
                """)).GetBoolean());
        }
    }

    [Fact]
    public async Task UnsupportedPopupNavigationIsRejectedBeforeChangingItsDocument()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><h1>Opener</h1>"
        });
        Assert.True((await session.EvaluateAsync("""
            (()=>{
                const popup=open();const original=popup.document;
                if(popup.location!==popup.document.location || popup.document.location.href!=='about:blank')return false;
                const operations=[()=>popup.location='https://popup.example/next',()=>popup.location.href='https://popup.example/next',
                    ()=>popup.location.assign('https://popup.example/next'),()=>popup.location.replace('about:blank'),
                    ()=>popup.location.reload(),()=>popup.location.hash='next',()=>popup.document.location='https://popup.example/next',
                    ()=>popup.document.location.href='https://popup.example/next'];
                for(const operation of operations){
                    let rejected=false;try{operation();}catch(error){rejected=error.name==='NotSupportedError';}
                    if(!rejected || popup.document!==original || original.URL!=='about:blank')return false;
                }
                return true;
            })()
            """)).GetBoolean());
    }

    [Fact]
    public async Task RejectedOpeningDoesNotConsumeTheCreationBudget()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, MaxAuxiliaryWindows = 1,
            Html = "<!doctype html><h1>Opener</h1>"
        });
        Assert.True((await session.EvaluateAsync("""
            (()=>{
                const failures=[];
                for(const invoke of [()=>open('https://elsewhere.example/'),()=>open('','_self'),()=>open('','','width=100'),()=>open.call({})]) {
                    try {invoke();failures.push('none');}catch(error){failures.push(error.name);}
                }
                const popup=open();
                return failures.join(',')==='NotSupportedError,NotSupportedError,NotSupportedError,TypeError' && popup!==null && open()===null;
            })()
            """)).GetBoolean());
    }

    [Fact]
    public async Task NestedAuxiliaryWindowSurvivesItsOpenersClosure()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, MaxAuxiliaryWindows = 2,
            Html = "<!doctype html><body>Opener"
        });
        await session.ExecuteAsync("""
            window.popup=open();window.nested=null;
            addEventListener('message',event=>{if(event.data==='nested-ready')nested=event.source;});
            const script=popup.document.createElement('script');
            script.textContent=`setTimeout(()=>{
                const child=open('','nested');const childScript=child.document.createElement('script');
                childScript.textContent="const root=opener.opener;setInterval(()=>root.document.body.setAttribute('data-tick',String(Number(root.document.body.getAttribute('data-tick')||0)+1)),20);root.postMessage('nested-ready','https://popup.example');";
                child.document.body.append(childScript);
            },0);`;
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("nested!==null");
        await session.WaitForAsync("Number(document.body.getAttribute('data-tick'))>0");
        await session.ExecuteAsync("popup.close();window.beforeClose=Number(document.body.getAttribute('data-tick'));");
        await session.WaitForAsync("Number(document.body.getAttribute('data-tick'))>beforeClose");
        Assert.True((await session.EvaluateAsync("popup.closed && !nested.closed && nested.opener===popup && nested.parent===nested && open()===null")).GetBoolean());
        Assert.True((await session.EvaluateAsync("open('','nested')===nested && nested.opener===window")).GetBoolean());
        await session.ExecuteAsync("nested.close()");
    }

    [Fact]
    public async Task PopupCannotNavigateItselfThroughItsGlobalOrDocumentLocation()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body>Opener"
        });
        await session.ExecuteAsync("""
            window.popup=open();
            const script=popup.document.createElement('script');
            script.textContent=`
                const initial=document;
                const attempts=[()=>location='https://popup.example/next',()=>window.location.href='https://popup.example/next',
                    ()=>document.location='https://popup.example/next',()=>document.location.assign('https://popup.example/next'),()=>location.reload()];
                const results=attempts.map(attempt=>{try{attempt();return 'accepted';}catch(error){return error.name;}});
                opener.document.body.setAttribute('data-results',results.join(','));
                opener.document.body.setAttribute('data-intact',String(document===initial && document.URL==='about:blank' && location===document.location));`;
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-results')");
        Assert.Equal(string.Join(",", Enumerable.Repeat("NotSupportedError", 5)), (await session.EvaluateAsync("document.body.getAttribute('data-results')")).GetString());
        Assert.Equal("true", (await session.EvaluateAsync("document.body.getAttribute('data-intact')")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ClosingPopupFencesAlreadyQueuedCallbacks(bool timer)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body>Opener"
        });
        string callback = timer
            ? "document.open();setTimeout(()=>host.document.body.setAttribute('data-late','yes'),0);close();"
            : "queueMicrotask(()=>close());Promise.resolve().then(()=>host.document.body.setAttribute('data-late','yes'));";
        await session.ExecuteAsync("window.popup=open();const script=popup.document.createElement('script');script.textContent="
            + System.Text.Json.JsonSerializer.Serialize("const host=opener;" + callback)
            + ";popup.document.body.append(script);");
        await session.WaitForAsync("popup.closed");
        await session.ExecuteAsync("window.settled=false;setTimeout(()=>settled=true,80)");
        await session.WaitForAsync("settled");
        Assert.False((await session.EvaluateAsync("document.body.hasAttribute('data-late')")).GetBoolean());
    }

    [Fact]
    public async Task RemovingFrameDuringMicrotaskDrainFencesLaterReactions()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body>Parent"
        });
        await session.ExecuteAsync("""
            const frame=document.createElement('iframe');document.body.append(frame);
            const script=frame.contentDocument.createElement('script');
            script.textContent="const host=parent;queueMicrotask(()=>frameElement.remove());Promise.resolve().then(()=>host.document.body.setAttribute('data-late','yes'));";
            frame.contentDocument.body.append(script);
            """);
        await session.ExecuteAsync("window.settled=false;setTimeout(()=>settled=true,80)");
        await session.WaitForAsync("settled");
        Assert.False((await session.EvaluateAsync("document.body.hasAttribute('data-late')")).GetBoolean());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RetiringRealmDisconnectsItsObserversOnForeignDocuments(bool frame)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body><button id='foreign'>Foreign</button>"
        });
        await session.ExecuteAsync(frame
            ? "window.childHost=document.createElement('iframe');document.body.append(childHost);window.popup=childHost.contentWindow;"
            : "window.popup=open();");
        await session.ExecuteAsync("""
            const script=popup.document.createElement('script');
            script.textContent=`const host=opener||parent;
                const observer=new MutationObserver(()=>host.document.body.setAttribute('data-observed',String(Number(host.document.body.getAttribute('data-observed')||0)+1)));
                observer.observe(host.document.getElementById('foreign'),{attributes:true});
                host.document.body.setAttribute('data-ready','yes');`;
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-ready')");
        await session.ExecuteAsync("document.getElementById('foreign').setAttribute('data-value','before')");
        await session.WaitForAsync("document.body.getAttribute('data-observed')==='1'");
        await session.ExecuteAsync(frame ? "childHost.remove();" : "popup.close();");
        await session.ExecuteAsync("document.getElementById('foreign').setAttribute('data-value','after')");
        Assert.Equal("1", (await session.EvaluateAsync("document.body.getAttribute('data-observed')")).GetString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RetiringRealmRemovesItsListenersFromRetainedAndForeignTargets(bool frame)
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start,
            Html = "<!doctype html><body><button id='foreign'>Foreign</button>"
        });
        await session.ExecuteAsync(frame
            ? "window.childHost=document.createElement('iframe');document.body.append(childHost);window.popup=childHost.contentWindow;"
            : "window.popup=open();");
        await session.ExecuteAsync("""
            const script=popup.document.createElement('script');
            script.textContent=`
                const host=window.opener||window.parent;
                const button=document.createElement('button');button.id='owned';document.body.append(button);
                const count=()=>host.document.body.setAttribute('data-count',String(Number(host.document.body.getAttribute('data-count')||0)+1));
                button.addEventListener('click',count);button.onclick=count;
                host.document.getElementById('foreign').addEventListener('click',count);
                host.document.body.setAttribute('data-ready','yes');`;
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("document.body.hasAttribute('data-ready')");
        await session.ExecuteAsync("""
            window.retained=popup.document.getElementById('owned');
            retained.click();document.getElementById('foreign').click();
            window.beforeClose=document.body.getAttribute('data-count');
            """);
        await session.ExecuteAsync(frame ? "childHost.remove();" : "popup.close();");
        await session.ExecuteAsync("retained.click();document.getElementById('foreign').click();");
        Assert.True((await session.EvaluateAsync("beforeClose==='3' && document.body.getAttribute('data-count')===beforeClose")).GetBoolean());
    }

    [Fact]
    public async Task PopupCloseRetiresItsTimerAndTheLifetimeBudgetDoesNotReset()
    {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = Start, MaxAuxiliaryWindows = 1,
            Html = "<!doctype html><h1>Opener</h1>"
        });
        await session.ExecuteAsync("""
            window.popup=open();
            const script=popup.document.createElement('script');
            script.textContent="setInterval(()=>opener.document.body.setAttribute('data-tick',String(Number(opener.document.body.getAttribute('data-tick')||0)+1)),20)";
            popup.document.body.append(script);
            """);
        await session.WaitForAsync("Number(document.body.getAttribute('data-tick'))>0");
        await session.ExecuteAsync("""
            popup.close();window.lastTick=document.body.getAttribute('data-tick');
            window.settled=false;setTimeout(()=>settled=true,80);
            """);
        await session.WaitForAsync("settled");
        Assert.True((await session.EvaluateAsync("popup.closed && document.body.getAttribute('data-tick')===lastTick && open()===null")).GetBoolean());
    }
}
