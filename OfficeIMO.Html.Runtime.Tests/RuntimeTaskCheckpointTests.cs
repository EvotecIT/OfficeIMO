using System.Text;
using OfficeIMO.Html.Providers;
using OfficeIMO.Html.Runtime;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class RuntimeTaskCheckpointTests {
    private static HtmlProcessRuntimeProvider Runtime() => new(
        Path.Combine(AppContext.BaseDirectory, "RuntimeWorker", "OfficeIMO.Html.Runtime.Worker.dll"), AngleSharpDomServices.Instance);

    [Theory]
    [InlineData("setTimeout")]
    [InlineData("setInterval")]
    public async Task NativeTimerJobsFinishBeforeTheFollowingTimerWithoutHostPolling(string timer) {
        await ObserveNativeTasksAsync("""
            window.order=[];
            const timer=TIMER(()=>{
                clearInterval(timer);
                order.push('timer');
                Promise.resolve().then(()=>{order.push('promise');queueMicrotask(()=>order.push('nested'))});
                queueMicrotask(()=>order.push('microtask'));
                setTimeout(()=>fetch('/result',{method:'POST',body:order.join(',')}),0);
            },0);
            """.Replace("TIMER", timer), (_, result) => {
                Assert.Equal("timer,promise,microtask,nested", result);
                return Task.CompletedTask;
            });
    }

    [Fact]
    public async Task ARejectionHandledByALaterTimerStillFailsTheSessionAtTheEarlierTaskBoundary() {
        await ObserveNativeTasksAsync("""
            setTimeout(()=>{
                const rejected=Promise.reject(new Error('late timer rejection'));
                setTimeout(()=>{
                    rejected.catch(()=>{});
                    fetch('/result',{method:'POST',body:'later task'});
                },0);
            },0);
            """, async (session, result) => {
                Assert.Equal("later task", result);
                var error = await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
                Assert.Contains("late timer rejection", error.Message);
            });
    }

    [Fact]
    public async Task NonterminatingIdleTimerMicrotaskIsInterruptedByTheNextCommandDeadline() {
        await ObserveNativeTasksAsync("""
            setTimeout(()=>{
                fetch('/result',{method:'POST',body:'entered'});
                queueMicrotask(()=>{while(true){}});
            },0);
            """, async (session, result) => {
                Assert.Equal("entered", result);
                await Assert.ThrowsAsync<TimeoutException>(() => session.CaptureAsync());
                await Assert.ThrowsAsync<HtmlScriptRuntimeException>(() => session.CaptureAsync());
            });
    }

    private static async Task ObserveNativeTasksAsync(string script, Func<IHtmlRuntimeSession, string, Task> verify) {
        var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var observed = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
        await using var server = new RuntimeHttpFixture((_, _) => Task.FromResult(RuntimeHttpFixture.Reply.Text("ok")));
        server.RespondToRequest = async (request, token) => {
            if (request.Path == "/start") await release.Task.WaitAsync(token);
            else if (request.Path == "/result") observed.TrySetResult(Encoding.UTF8.GetString(request.Body));
            return RuntimeHttpFixture.Reply.Text("ok");
        };
        await using var session = await Runtime().OpenTrustedAsync(new() {
            // The same timeout covers worker startup and the later command. Keep
            // normal startup headroom when the full suite starts workers concurrently.
            DocumentUrl = server.Origin, ResourcePolicy = new() { AllowNetwork = true }, Timeout = TimeSpan.FromSeconds(10)
        });
        await session.ExecuteAsync("fetch('/start').then(()=>{" + script + "\n})");
        // Release the real network boundary only after command completion. Observe the
        // later timer through HTTP, so no session polling can accidentally drain jobs.
        release.SetResult();
        await verify(session, await observed.Task.WaitAsync(TimeSpan.FromSeconds(10)));
    }
}
