using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeResizableHistoryStateTests {
    private static readonly Uri First = new("https://buffers.example/first");
    private static readonly Uri Second = new(First, "/second");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task ResizableGraphsRetainAliasesAndViewModesAcrossReloadAndTraversal() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = First, Html = "<h1>Buffer history</h1>",
            Resources = new[] { HtmlRuntimeResource.FromText(Second, "<h1>Other page</h1>", "text/html") }
        });
        await session.ExecuteAsync("""
            const buffer=new ArrayBuffer(8,{maxByteLength:24});
            const tracking=new Uint16Array(buffer,2), fixed=new Uint16Array(buffer,2,3);
            tracking[0]=73;
            const state={buffer,tracking,fixed,data:new DataView(buffer,2),fixedData:new DataView(buffer,2,6),
                empty:new Uint8Array(buffer,8,0),tail:new Uint8Array(buffer,8)};
            state.self=state;state.again=tracking;
            history.replaceState(state,'');tracking[0]=99;buffer.resize(16);
            """);
        const string valid = """
            (()=>{const s=history.state;
                const saved=s.self===s && s.again===s.tracking && s.buffer instanceof ArrayBuffer &&
                    s.buffer.resizable && s.buffer.maxByteLength===24 && s.buffer.byteLength===8 &&
                    s.tracking instanceof Uint16Array && s.data instanceof DataView &&
                    s.tracking.buffer===s.buffer && s.fixed.buffer===s.buffer && s.data.buffer===s.buffer && s.fixedData.buffer===s.buffer &&
                    s.tracking[0]===73 && s.tracking.byteOffset===2 && s.empty.byteOffset===8;
                s.buffer.resize(16);
                const grown=s.tracking.length===7 && s.fixed.length===3 && s.data.byteLength===14 && s.fixedData.byteLength===6 && s.tail.length===8 && s.empty.length===0;
                s.buffer.resize(1);s.buffer.resize(8);
                return saved && grown && s.tracking.length===3 && s.fixed.length===3 && s.data.byteLength===6;
            })()
            """;
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean());
        await session.ReloadAsync();
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean());
        await session.NavigateAsync(Second);
        await session.ExecuteAsync("history.back()");
        await session.WaitForAsync("location.pathname==='/first'");
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean());
    }

    [Theory]
    [InlineData("pushState")]
    [InlineData("replaceState")]
    public async Task InvalidViewsAndSharedBuffersFailWithoutChangingHistory(string operation) {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First });
        await session.ExecuteAsync($$"""
            history.replaceState({saved:true},'');window.previous=history.state;window.failures=[];
            const buffer=new ArrayBuffer(8,{maxByteLength:16});
            const fixed=new Uint8Array(buffer,2,6),tracking=new Uint8Array(buffer,2),data=new DataView(buffer,2);
            buffer.resize(1);
            const detached=new ArrayBuffer(8,{maxByteLength:16}),detachedView=new Uint8Array(detached);detached.transfer();
            for(const value of [fixed,tracking,data,detached,detachedView,new SharedArrayBuffer(8),new Uint8Array(new SharedArrayBuffer(8))]) {
                try{history.{{operation}}(value,'','/invalid');failures.push('accepted')}catch(e){failures.push(e.name)}
            }
            """);
        Assert.Equal(string.Join(',', Enumerable.Repeat("DataCloneError", 7)), (await session.EvaluateAsync("failures.join(',')")).GetString());
        Assert.True((await session.EvaluateAsync("history.state===previous && history.length===1 && location.pathname==='/first'")).GetBoolean());
    }

    [Fact]
    public async Task BufferBudgetsChargeBytesOncePerGraphAndRejectOversizedStateAtomically() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First, MaxHistoryStateBytes = 512 });
        await session.ExecuteAsync("""
            const small=new ArrayBuffer(128,{maxByteLength:1024});history.replaceState([small,small],'');
            window.previous=history.state;window.failure='';
            try{history.pushState(new ArrayBuffer(1024,{maxByteLength:2048}),'','/large')}catch(e){failure=e.name}
            """);
        Assert.True((await session.EvaluateAsync("failure==='DataCloneError' && history.state===previous && previous[0]===previous[1] && previous[0].maxByteLength===1024 && history.length===1 && location.pathname==='/first'")).GetBoolean());
    }

    [Fact]
    public async Task FrameMessagesPreserveNativeBufferMetadataAndRejectOutOfBoundsViews() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = First,
            Html = "<script>onmessage=e=>{window.received=e.data;history.replaceState(e.data,'')}</script><iframe src='/child'></iframe>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(First, "/child"), """
                <script>
                const buffer=new ArrayBuffer(8,{maxByteLength:16}),tracking=new Uint8Array(buffer,2),fixed=new Uint8Array(buffer,2,6),data=new DataView(buffer,2);
                tracking[0]=42;
                const state={buffer,tracking,fixed,data};state.self=state;
                Object.defineProperty(buffer,'maxByteLength',{get(){throw new Error('authored buffer getter')}});
                Object.defineProperty(tracking,'length',{get(){throw new Error('authored view getter')}});
                parent.postMessage(state,'*');buffer.resize(1);
                try{parent.postMessage(fixed,'*')}catch(e){parent.document.body.dataset.failure=e.name}
                </script>
                """, "text/html") }
        });
        await session.WaitForAsync("!!window.received");
        Assert.True((await session.EvaluateAsync("""
            (()=>{const s=received,h=history.state;
            const valid=s.self===s && h.self===h && s.buffer!==h.buffer && s.tracking.buffer===s.buffer && s.data.buffer===s.buffer &&
                s.buffer.resizable && s.buffer.maxByteLength===16 && s.tracking[0]===42 && h.tracking[0]===42;
            s.buffer.resize(16);h.buffer.resize(12);
            return valid && s.tracking.length===14 && s.fixed.length===6 && s.data.byteLength===14 && h.tracking.length===10 && h.fixed.length===6})()
            """)).GetBoolean());
        Assert.Equal("DataCloneError", (await session.EvaluateAsync("document.body.dataset.failure")).GetString());
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FrameBufferBytesIgnoreInheritedArrayAccessors(bool resizable) {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = First,
            Html = "<script>onmessage=e=>{window.lastByte=new Uint8Array(e.data)[7]}</script><iframe src='/child'></iframe>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(First, "/child"), $$$"""
                <script>
                const buffer=new ArrayBuffer(8,{{{(resizable ? "{maxByteLength:16}" : "undefined")}}});
                new Uint8Array(buffer)[7]=42;let calls=0;
                Object.defineProperty(Array.prototype,'7',{configurable:true,get(){return 99},set(){calls++}});
                try{parent.postMessage(buffer,'*')}finally{delete Array.prototype[7]}
                parent.document.body.dataset.calls=String(calls);
                </script>
                """, "text/html") }
        });
        await session.WaitForAsync("window.lastByte!==undefined");
        Assert.Equal(42, (await session.EvaluateAsync("lastByte")).GetInt32());
        Assert.Equal("0", (await session.EvaluateAsync("document.body.dataset.calls")).GetString());
    }

}
