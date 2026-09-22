using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeDatasetTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Fact]
    public async Task DatasetReflectsLiveAttributesAndDeletionInFrozenCapture() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest {
            Html = "<body><p id='status' data-foo-bar='first' data-to-string='text'>Waiting</p></body>"
        });
        var result = await session.EvaluateAsync("""
            (() => {
                const element = document.querySelector('#status');
                const data = element.dataset;
                const initial = [data.fooBar, data.toString, Object.keys(data), 'missing' in data, Object.hasOwn(data,'missing')];
                element.setAttribute('data-foo-bar', 'second');
                const current = data.fooBar;
                const unsupportedDelete = delete data['foo-bar'];
                const retained = data.fooBar;
                const removed = delete data.fooBar;
                data.constructor = 'report';
                Object.defineProperty(data, 'runCount', { value: 2 });
                const descriptor = Object.getOwnPropertyDescriptor(data, 'runCount');
                element.textContent = Object.entries(data).map(([key,value]) => key + '=' + value).join(', ');
                return { initial, current, unsupportedDelete, retained, removed,
                    sameObject: data === element.dataset,
                    keys: Object.keys(data), descriptor,
                    copied: {...data}, text: element.textContent };
            })()
            """);
        Assert.Equal("[\"first\",\"text\",[\"fooBar\",\"toString\"],false,false]", result.GetProperty("initial").GetRawText());
        Assert.Equal("second", result.GetProperty("current").GetString());
        Assert.True(result.GetProperty("unsupportedDelete").GetBoolean());
        Assert.Equal("second", result.GetProperty("retained").GetString());
        Assert.True(result.GetProperty("removed").GetBoolean());
        Assert.True(result.GetProperty("sameObject").GetBoolean());
        Assert.Equal("[\"toString\",\"constructor\",\"runCount\"]", result.GetProperty("keys").GetRawText());
        Assert.Equal("{\"value\":\"2\",\"writable\":true,\"enumerable\":true,\"configurable\":true}", result.GetProperty("descriptor").GetRawText());
        Assert.Equal("2", result.GetProperty("copied").GetProperty("runCount").GetString());
        var capture = await session.CaptureAsync();
        var status = capture.Document.QuerySelector("#status")!;
        Assert.Null(status.GetAttribute("data-foo-bar"));
        Assert.Equal("2", status.GetAttribute("data-run-count"));
        Assert.Equal("toString=text, constructor=report, runCount=2", status.TextContent);
    }

    [Fact]
    public async Task DatasetConvertsValuesAndReportsCatchableDomErrors() {
        await using var session = await Runtime().OpenTrustedAsync(new HtmlScriptRequest { Html = "<body></body>" });
        var result = await session.EvaluateAsync("""
            (() => {
                const data = document.body.dataset;
                data.emptyValue = null;
                data.missingValue = undefined;
                const errors = [];
                for (const key of ['bad-key','bad key']) {
                    try { data[key] = 'invalid'; }
                    catch (error) { errors.push([error.name,error.code,typeof error.message]); }
                }
                try { data.symbolValue = Symbol(); } catch(error) { errors.push([error.name]); }
                const symbol = Symbol('metadata');
                data[symbol] = 42;
                const symbolValue = data[symbol];
                const deleted = delete data[symbol];
                return { values:[data.emptyValue,data.missingValue], errors, symbolValue, deleted,
                    keys:Object.keys(data), accessor:Reflect.defineProperty(data,'accessor',{get:()=>1}),
                    prevented:Reflect.preventExtensions(data) };
            })()
            """);
        Assert.Equal("[\"null\",\"undefined\"]", result.GetProperty("values").GetRawText());
        Assert.Equal("[[\"SyntaxError\",12,\"string\"],[\"InvalidCharacterError\",5,\"string\"],[\"TypeError\"]]", result.GetProperty("errors").GetRawText());
        Assert.Equal(42, result.GetProperty("symbolValue").GetInt32());
        Assert.True(result.GetProperty("deleted").GetBoolean());
        Assert.Equal("[\"emptyValue\",\"missingValue\"]", result.GetProperty("keys").GetRawText());
        Assert.False(result.GetProperty("accessor").GetBoolean());
        Assert.False(result.GetProperty("prevented").GetBoolean());
    }
}
