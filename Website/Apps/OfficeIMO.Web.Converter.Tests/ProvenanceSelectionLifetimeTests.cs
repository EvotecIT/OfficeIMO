using System.Net;
using System.Reflection;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Components;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class ProvenanceSelectionLifetimeTests {
    [Theory]
    [InlineData(false, "clear")]
    [InlineData(true, "clear")]
    [InlineData(false, "tool")]
    [InlineData(true, "tool")]
    [InlineData(false, "dispose")]
    [InlineData(true, "dispose")]
    [InlineData(false, "unchanged")]
    [InlineData(true, "unchanged")]
    public async Task PendingSelectionCommitsOnlyWhileItsWorkspaceStillOwnsIt(bool sample, string action) {
        var session = new BrowserDocumentSession();
        var original = new SelectedDocument("original.png", ".png", "PNG", 1, [1]);
        session.Open([original]);
        var js = new DelayedRevocation();
        using var http = new HttpClient(new SampleHandler()) { BaseAddress = new Uri("https://example.test/") };
        var component = new ProvenanceWorkbench();
        Set(component, "Session", session); Set(component, "Http", http);
        Set(component, "_interop", new ConverterInterop(js)); Set(component, "_file", original);
        Set(component, "_outputUrl", "blob:previous");
        var method = typeof(ProvenanceWorkbench).GetMethod(sample ? "LoadSampleAsync" : "OpenAsync", BindingFlags.Instance | BindingFlags.NonPublic)!;
        var pending = (Task)method.Invoke(component, sample ? null : [new InputFileChangeEventArgs([new UploadedImage()])])!;
        await js.Entered.Task.WaitAsync(TimeSpan.FromSeconds(5));
        if (action == "clear") session.Clear();
        if (action == "tool") session.ChangeTool();
        if (action == "dispose") await component.DisposeAsync();
        js.Release.SetResult();
        await pending.WaitAsync(TimeSpan.FromSeconds(5));
        if (action == "clear") Assert.Empty(session.Current);
        else if (action == "unchanged") Assert.Equal(sample ? "provenance-demo.png" : "replacement.png", Assert.Single(session.Current).Name);
        else Assert.Same(original, Assert.Single(session.Current));
        await component.DisposeAsync();
    }

    private static void Set(object target, string name, object value) {
        const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic;
        var property = target.GetType().GetProperty(name, flags);
        if (property is not null) property.SetValue(target, value);
        else target.GetType().GetField(name, flags)!.SetValue(target, value);
    }
    private sealed class DelayedRevocation : IJSRuntime {
        internal TaskCompletionSource Entered { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args) => InvokeAsync<TValue>(identifier, default, args);
        public async ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args) {
            Assert.Equal("URL.revokeObjectURL", identifier);
            Entered.TrySetResult(); await Release.Task;
            return default!;
        }
    }
    private sealed class SampleHandler : HttpMessageHandler {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
            Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent([2]) });
    }
    private sealed class UploadedImage : IBrowserFile {
        public string Name => "replacement.png";
        public DateTimeOffset LastModified => DateTimeOffset.UnixEpoch;
        public long Size => 1;
        public string ContentType => "image/png";
        public Stream OpenReadStream(long maxAllowedSize = 512000, CancellationToken cancellationToken = default) => new MemoryStream([2]);
    }
}
