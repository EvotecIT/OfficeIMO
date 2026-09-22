using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeHistoryStateTests {
    private static readonly Uri First = new("https://state.example/first");
    private static readonly Uri Second = new(First, "/second");
    private static HtmlProcessRuntimeProvider Runtime() => new(Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task NativeGraphsSurviveReloadAndCrossDocumentTraversalInTheNewRealm() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            Profile = HtmlRuntimeProfile.WebApplicationV1, DocumentUrl = First,
            Html = "<body><h1>State report</h1></body>",
            Resources = new[] { HtmlRuntimeResource.FromText(Second, "<h1>Other page</h1>", "text/html") }
        });
        await session.ExecuteAsync("""
            const key={id:42}, buffer=new ArrayBuffer(8), view=new Uint16Array(buffer,2,2);
            view[0]=73;
            const error=new TypeError('Saved error',{cause:key});error.stack='Saved error trace';
            const state={key,buffer,view,data:new DataView(buffer,2,4),error,
                boxes:[new Boolean(false),new Number(-0),new String('saved'),Object(123n)],
                invalidDate:new Date(NaN),date:new Date(123),regex:/route/gi,
                map:new Map([[key,error]]),set:new Set([key,error]),values:[undefined,NaN,Infinity,-Infinity,-0]};
            state.self=state;error.cause=state;state.sparse=[,key];
            history.replaceState(state,'');view[0]=99;key.id=99;
            """);
        const string valid = """
            (()=>{const s=history.state;return s.self===s && s.key.id===42 &&
            s.buffer instanceof ArrayBuffer && s.view instanceof Uint16Array && s.data instanceof DataView &&
            s.view.buffer===s.buffer && s.data.buffer===s.buffer && s.view[0]===73 && s.view.byteOffset===2 &&
            s.error instanceof TypeError && s.error.message==='Saved error' && s.error.stack==='Saved error trace' && s.error.cause===s &&
            s.boxes[0] instanceof Boolean && s.boxes[0].valueOf()===false && s.boxes[1] instanceof Number && Object.is(s.boxes[1].valueOf(),-0) &&
            s.boxes[2] instanceof String && s.boxes[2].valueOf()==='saved' && Object.getPrototypeOf(s.boxes[3])===BigInt.prototype && s.boxes[3].valueOf()===123n &&
            s.invalidDate instanceof Date && Number.isNaN(s.invalidDate.getTime()) && s.date.getTime()===123 && s.regex instanceof RegExp && s.regex.flags==='gi' &&
            s.map instanceof Map && s.map.get(s.key)===s.error && s.set instanceof Set && s.set.has(s.key) && s.set.has(s.error) &&
            s.values[0]===undefined && Number.isNaN(s.values[1]) && s.values[2]===Infinity && s.values[3]===-Infinity && Object.is(s.values[4],-0) && !(0 in s.sparse) && s.sparse[1]===s.key})()
            """;
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean(), (await session.EvaluateAsync("({stack:history.state.error.stack,boxTypes:history.state.boxes.map(x=>Object.prototype.toString.call(x)),name:history.state.error.name,own:Object.getOwnPropertyNames(history.state.error)})")).ToString());
        await session.ReloadAsync();
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean(), (await session.EvaluateAsync("({stack:history.state.error.stack,boxTypes:history.state.boxes.map(x=>Object.prototype.toString.call(x)),name:history.state.error.name,own:Object.getOwnPropertyNames(history.state.error)})")).ToString());
        await session.NavigateAsync(Second);
        await session.ExecuteAsync("history.back()");
        await session.WaitForAsync("location.pathname==='/first'");
        Assert.True((await session.EvaluateAsync(valid)).GetBoolean(), (await session.EvaluateAsync("({stack:history.state.error.stack,boxTypes:history.state.boxes.map(x=>Object.prototype.toString.call(x)),name:history.state.error.name,own:Object.getOwnPropertyNames(history.state.error)})")).ToString());
    }

    [Theory]
    [InlineData("pushState")]
    [InlineData("replaceState")]
    public async Task ErrorMessageSymbolThrowsBeforeChangingHistory(string operation) {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First });
        await session.ExecuteAsync("history.replaceState({saved:true},'');window.previous=history.state;const error=new Error();error.message=Symbol('bad');window.failure='';try{history." + operation + "(error,'','/changed')}catch(e){failure=e.name}");
        Assert.Equal("TypeError", (await session.EvaluateAsync("failure")).GetString());
        Assert.True((await session.EvaluateAsync("history.state===previous && history.state.saved && history.length===1 && location.pathname==='/first'")).GetBoolean());
    }

    [Fact]
    public async Task ErrorAccessorFieldsAreNotInvokedAndCustomNamesUseError() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First });
        await session.ExecuteAsync("""
            const error=new Error();error.name='ApplicationError';
            Object.defineProperty(error,'message',{get(){throw new Error('message getter')}});
            Object.defineProperty(error,'cause',{get(){throw new Error('cause getter')}});
            error.stack='Application trace';history.replaceState(error,'');
            """);
        Assert.True((await session.EvaluateAsync("history.state instanceof Error && history.state.name==='Error' && !Object.hasOwn(history.state,'message') && !Object.hasOwn(history.state,'cause') && history.state.stack==='Application trace'")).GetBoolean(), (await session.EvaluateAsync("({stack:history.state.stack,name:history.state.name,own:Object.getOwnPropertyNames(history.state)})")).ToString());
    }
    [Fact]
    public async Task ErrorStackBudgetAndThrowingNameLeaveThePreviousStateIntact() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First, MaxHistoryStateBytes = 512 });
        await session.ExecuteAsync("""
            history.replaceState({saved:true},'');window.previous=history.state;window.failures=[];
            const large=new Error('Large');large.stack='trace'.repeat(200);
            try{history.pushState(large,'','/large')}catch(e){failures.push(e.name)}
            const throwing=new Error();const sentinel=new RangeError('name');
            Object.defineProperty(throwing,'name',{get(){throw sentinel}});
            try{history.replaceState(throwing,'','/throw')}catch(e){failures.push(e===sentinel?'sentinel':'wrong')}
            """);
        Assert.Equal("DataCloneError,sentinel", (await session.EvaluateAsync("failures.join(',')")).GetString());
        Assert.True((await session.EvaluateAsync("history.state===previous && history.length===1 && location.pathname==='/first'")).GetBoolean());
    }

    [Fact]
    public async Task CapturedNativeStackSurvivesHistoryCopiesWithoutCallingAStackAccessor() {
        await using var session = await Runtime().OpenTrustedAsync(new() { DocumentUrl = First });
        await session.ExecuteAsync("""
            function originalFailure(){return new RangeError('Original failure')}
            const error=originalFailure();window.originalTrace=error.stack;
            Object.defineProperty(error,'stack',{get(){throw new Error('stack accessor')}});
            history.replaceState(error,'');
            """);
        Assert.True((await session.EvaluateAsync("history.state instanceof RangeError && originalTrace.length>0 && history.state.stack===originalTrace")).GetBoolean());
    }

    [Fact]
    public async Task FrameErrorGraphsUseTheSamePayloadRulesAsHistory() {
        await using var session = await Runtime().OpenTrustedAsync(new() {
            DocumentUrl = First,
            Html = "<body><script>onmessage=e=>{window.received=e.data;history.replaceState(e.data,'')}</script><iframe src='/child'></iframe></body>",
            Resources = new[] { HtmlRuntimeResource.FromText(new Uri(First, "/child"), """
                <script>
                const error=new TypeError('Child error');error.stack='Child trace';error.cause=error;
                parent.postMessage(error,'*');
                const invalid=new Error();invalid.message=Symbol('invalid');
                try{parent.postMessage(invalid,'*')}catch(e){parent.document.body.dataset.failure=e.name}
                </script>
                """, "text/html") }
        });
        await session.WaitForAsync("!!window.received");
        Assert.True((await session.EvaluateAsync("received instanceof TypeError && received.cause===received && received.stack==='Child trace' && history.state!==received && history.state.cause===history.state && history.state.stack==='Child trace'")).GetBoolean());
        Assert.Equal("TypeError", (await session.EvaluateAsync("document.body.dataset.failure")).GetString());
    }

}
