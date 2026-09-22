using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeSharedStorageTests {
    [Theory]
    [InlineData("localStorage")]
    [InlineData("sessionStorage")]
    public async Task SameOriginFramesObserveCurrentStorageAndShareQuota(string area) {
        var runtime = new HtmlProcessRuntimeProvider(
            Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);
        await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
            DocumentUrl = new Uri("https://app.example/index.html"),
            MaxStorageCharacters = 8,
            Html = "<body><iframe src='/child.html'></iframe></body>",
            Resources = [HtmlRuntimeResource.FromText(new Uri("https://app.example/child.html"), $$"""
                <body><script>
                addEventListener('message', () => {
                    const store = {{area}};
                    parent.document.body.dataset.seen = JSON.stringify([store.getItem('key'), store.key(0), store.length, Object.keys(store), store.key]);
                    try { store.other = 'too large'; } catch (error) { parent.document.body.dataset.quota = error.name; }
                    store.removeItem('key');
                    store.fresh = 'yes';
                    parent.document.body.dataset.done = 'yes';
                });
                parent.document.body.dataset.ready = 'yes';
                </script></body>
                """, "text/html")]
        });
        await session.WaitForAsync("document.body.dataset.ready==='yes'");
        await session.ExecuteAsync(area + ".setItem('key','value');document.querySelector('iframe').contentWindow.postMessage('read','*')");
        await session.WaitForAsync("document.body.dataset.done==='yes'");
        var result = await session.EvaluateAsync("({seen:JSON.parse(document.body.dataset.seen),quota:document.body.dataset.quota,value:" + area + ".fresh,removed:" + area + ".getItem('key')})");
        Assert.Equal("value", result.GetProperty("seen")[0].GetString());
        Assert.Equal("key", result.GetProperty("seen")[1].GetString());
        Assert.Equal(1, result.GetProperty("seen")[2].GetInt32());
        Assert.Equal("QuotaExceededError", result.GetProperty("quota").GetString());
        Assert.Equal("yes", result.GetProperty("value").GetString());
        Assert.Null(result.GetProperty("removed").GetString());
    }
}
